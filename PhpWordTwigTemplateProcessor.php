<?php

namespace App\Helpers;

use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\Shared\ZipArchive;
use PhpOffice\PhpWord\TemplateProcessor;
use Twig\Environment;
use Twig\Error\LoaderError;
use Twig\Error\SyntaxError;
use Twig\Loader\ArrayLoader;

/**
 * Расширение шаблонизатора PhpWord, которое использует шаблонизатор Twig для динамического заполнения документа данными.
 * @see TemplateProcessor Исходный шаблонизатор PhpWord
 */
class PhpWordTwigTemplateProcessor extends TemplateProcessor
{
    private array $TAGS_TO_MOVE_OUT = [
        // Выносим Twig-блоки "{%tr ... %}" из строк таблицы
        'tr',
        // Выносим Twig-блоки "{%p ... %}" из абзацев
        'p',
    ];

    private readonly string $TAGS_TO_MOVE_OUT_REGEX;

    private readonly string $IMAGE_PREFIX = ' %%% ';

    private readonly string $IMAGE_LINK_END = ' %%%%% ';

    private readonly string $IMAGE_POSTFIX = ' %%% ';

    /**
     * @throws CopyFileException
     * @throws CreateTemporaryFileException
     *
     * @noinspection PhpMissingParentConstructorInspection
     * @see          TemplateProcessor::__construct() Полностью идентичен, но нужен, для внутреннего вызова своего метода fixBrokenMacros
     */
    public function __construct($documentTemplate)
    {
        // Temporary document filename initialization
        $this->tempDocumentFilename = tempnam(Settings::getTempDir(), 'PhpWord');
        if (false === $this->tempDocumentFilename) {
            throw new CreateTemporaryFileException();
        }

        // Template file cloning
        if (false === copy($documentTemplate, $this->tempDocumentFilename)) {
            throw new CopyFileException($documentTemplate, $this->tempDocumentFilename);
        }

        // Temporary document content extraction
        $this->zipClass = new ZipArchive();
        $this->zipClass->open($this->tempDocumentFilename);

        $index = 1;
        while (false !== $this->zipClass->locateName($this->getHeaderName($index))) {
            $this->tempDocumentHeaders[$index] = $this->readPartWithRels($this->getHeaderName($index));
            ++$index;
        }

        $index = 1;
        while (false !== $this->zipClass->locateName($this->getFooterName($index))) {
            $this->tempDocumentFooters[$index] = $this->readPartWithRels($this->getFooterName($index));
            ++$index;
        }

        $this->tempDocumentMainPart = $this->readPartWithRels($this->getMainPartName());
        $this->tempDocumentSettingsPart = $this->readPartWithRels($this->getSettingsPartName());
        $this->tempDocumentContentTypes = $this->zipClass->getFromName($this->getDocumentContentTypesName());

        $this->TAGS_TO_MOVE_OUT_REGEX = '(' . implode('|', $this->TAGS_TO_MOVE_OUT) . ')';
    }

    /**
     * @see TemplateProcessor::readPartWithRels() Полностью идентичен, но нужен, для внутреннего вызова своего метода fixBrokenMacros
     */
    protected function readPartWithRels($fileName): array|string
    {
        $relsFileName = $this->getRelationsName($fileName);
        $partRelations = $this->zipClass->getFromName($relsFileName);
        if ($partRelations !== false) {
            $this->tempDocumentRelations[$fileName] = $partRelations;
        }

        return $this->fixBrokenMacros($this->zipClass->getFromName($fileName));
    }

    /**
     * Убирает разметку XML внутри синтаксиса Twig, чтобы преобразовать XML в полноценный шаблон для Twig.
     * Применяется для блоков "{{ ... }}", "{% ... %}", "{# ... #}".
     * @see TemplateProcessor::fixBrokenMacros() Взят за основу. Вместо синтаксиса "${ ... }" используем синтаксис Twig.
     */
    protected function fixBrokenMacros($documentPart): array|string
    {
        // Так как при вводе в Word, скобки могут оказаться в разных тегах "<w:t>", учитываем это.
        // Также учитываем, что Twig-макрос может содержать специальные символы "{", "}" и "%".
        //
        // Группы (указаны для понимания):
        // - Открывающая скобка Twig-блока
        // - Возможный XML между открывающими символами Twig-блока
        // - Второй символ Twig-блока
        //   - "{" - выражение
        //   - "%" - блок кода
        //   - "#" - комментарий
        // - Содержимое Twig-блока
        // - Второй символ Twig-блока
        //   - "}" - выражение
        //   - "%" - блок кода
        //   - "#" - комментарий
        // - Возможный XML между закрывающими символами Twig-блока
        // - Закрывающая скобка Twig-блока
        return preg_replace_callback(
            '/(\{)(<\/w:t>.*?)?([{%#])(.*?)([#%}])(<\/w:t>.*?)?(})/',
            function (array $match): array|string {
                // Удаляем весь XML между скобок Twig-блока
                // $match[0] - всё подходящее регулярное выражение
                $removed_tags = strip_tags($match[0]);
                // Заменяем кавычки, чтобы использовать любые (в основном, пригодится при вводе кавычек в Word)
                // (Тут не используем функцию preg_replace, так как она может сломать кодировку)
                return str_replace(
                    ["‘", "’", "“", "”"],
                    ["'", "'", '"', '"'],
                    $removed_tags
                );
            },
            $documentPart,
        ) ?? $documentPart;
    }

    /**
     * Выносит Twig-блок из указанного XML-тега в разметке Word.
     * @param  string $documentPart Содержимое документа в виде XML
     * @param  string $tag_name     Название XML-тега в разметке Word.
     *                              Например, для тега "<w:tr>" значение этой переменной должно быть "tr".
     *                              При этом в самом документе, Twig-блок должен:
     *                              - иметь это же название сразу после открытия;
     *                              - иметь после названия хотя бы один пробел.
     *                              Для примера с "<w:tr>" Twig-блок должен иметь формат "{%tr ... %}".
     * @return string Новое содержимое документа в виде XML
     */
    private function moveOutMacrosFromTag(string $documentPart, string $tag_name): string
    {
        // XML-тег для ячейки таблицы
        $regex = $this->getRegexToFindXMLTags($tag_name);
        // Находим все XML-теги в разметке Word с указанным названием
        return preg_replace_callback(
            $regex,
            function (array $match) use ($tag_name): string {
                $xml_tag_content = $match[0];

                // Если содержимое XML-тег (включая сам XML-тег) содержит блок Twig, то:
                // - Очищаем весь XML, оставляя только блок Twig;
                // - Удаляем название тега из Twig-блока.
                //   Например, при названии тега "tr": "{%tr for i in arr %}" станет "{% for i in arr %}"
                $is_match = preg_match('/.*(\{[{%#]\s*)' . $tag_name . '(\s+.*?[#%}]}).*/', $xml_tag_content, $inner_matches);
                if ($is_match === 1) {
                    return $inner_matches[1] . $inner_matches[2];
                }

                // Если тег не содержит блок Twig - оставляем всё как есть
                return $xml_tag_content;
            },
            $documentPart,
        ) ?? $documentPart;
    }

    private function getRegexToFindXMLTags(string $tag_name): string
    {
        return '/<w:' . $tag_name . '( .*?)?>.+?<\/w:' . $tag_name . '>/';
    }

    private function combineTableRows(string $documentPart): string
    {
        // XML-тег для ячейки таблицы
        $regex_tc = $this->getRegexToFindXMLTags('tc');
        // XML-тег для настроек ячейки таблицы
        $regex_tcPr = $this->getRegexToFindXMLTags('tcPr');
        // XML-тег для настроек ячейки таблицы - Настройка вертикального объединения
        $regex_vMerge = $this->getRegexToFindXMLTags('vMerge');

        // Находим все XML-теги в разметке Word с указанным названием
        return preg_replace_callback(
            $regex_tc,
            function (array $match) use ($regex_vMerge, $regex_tcPr): array|string|null {
                $xml_tag_content = $match[0];

                // Если ячейка таблицы содержит Twig-блок "{% vm %}"
                $is_match = preg_match('/(.*)\{%\s*?vm\s*?%}(.*)/', $xml_tag_content, $inner_matches);
                if ($is_match === 1) {
                    // Удаляем тег "{% vm %}"
                    $xml_tag_content = $inner_matches[1] . $inner_matches[2];

                    return preg_replace_callback(
                        $regex_tcPr,
                        function (array $match) use ($regex_vMerge): string|array|null {
                            $tcPr_tag_content = $match[0];
                            // Удаляем существующую настройку vMerge (если есть)
                            $tcPr_tag_content = preg_replace_callback(
                                $regex_vMerge,
                                fn (): string => '',
                                $tcPr_tag_content,
                            );
                            // Вставляем новую настройку vMerge
                            return preg_replace('/<\/w:tcPr>/', '<w:vMerge w:val="{% if loop.first %}restart{% else %}continue{% endif %}"/></w:tcPr>', $tcPr_tag_content);
                        },
                        $xml_tag_content,
                    );
                }

                // Если ячейка таблицы не содержит Twig-блок "{% vm %}" - оставляем всё как есть
                return $xml_tag_content;
            },
            $documentPart,
        ) ?? $documentPart;
    }

    /**
     * Подготавливает теги "{% picture ... %} ... {% endpicture %}" для вставки изображений.
     * Между этих тегов в самом шаблоне Word необходимо разместить изображение, которое будет заменено.
     *
     * @param  string $documentPart Содержимое документа в виде XML
     * @return string Новое содержимое документа в виде XML
     */
    private function prepareImages(string $documentPart): string
    {
        // Заключаем само изображение в блок "{% if ... %}", проверяющий, что значение URI для изображения не пустое.
        // Если оно пустое - изображение не будет вставлено (ничего не выведется).

        // Находим теги "{% picture ... %}" и подготавливаем их
        $documentPart = preg_replace_callback(
            '/\{%\s*?(' . $this->TAGS_TO_MOVE_OUT_REGEX . '\s+?)?picture\s+?(.*?)\s*?%}/',
            function (array $match): string {
                $image_uri_variable = $match[3];
                return '{% if ' . $image_uri_variable . ' %}' . $this->IMAGE_PREFIX . '{{ ' . $image_uri_variable . '}}' . $this->IMAGE_LINK_END;
            },
            $documentPart,
        ) ?? $documentPart;

        // Находим теги "{% endpicture %}" и подготавливаем их
        return preg_replace_callback(
            '/\{%\s*?(' . $this->TAGS_TO_MOVE_OUT_REGEX . '\s+?)?endpicture\s+?%}/',
            fn (): string => $this->IMAGE_POSTFIX . '{% endif %}',
            $documentPart,
        ) ?? $documentPart;
    }

    /**
     * Вставляет изображения в подготовленные места для их вставки.
     * @param  string $documentPart Содержимое документа в виде XML
     * @return string Новое содержимое документа в виде XML
     */
    private function insertImages(string $documentPart): string
    {
        $rels = $this->zipClass->getFromName("word/_rels/document.xml.rels");

        // Находим все XML-теги в разметке Word с указанным названием
        return preg_replace_callback(
            '/' . $this->IMAGE_PREFIX . '(.+?)' . $this->IMAGE_LINK_END . '(.+?)' . $this->IMAGE_POSTFIX . '/',
            function (array $match) use ($rels): string {
                $image_uri = $match[1];
                $image_xml = $match[2];

                // Находим ID связки на изображение
                $image_rel_id = preg_replace('/.*r:embed="(.+?)".*/', '\1', $image_xml);

                // Находим по ID связки само изображение в архиве
                $image_path = preg_replace('/.*Relationship Id="' . $image_rel_id . '" Type=".+?" Target="(.+?)".*/s', '\1', $rels);

                // Скачивание изображения отдельным временным файлом
                $file_name_result = tempnam(sys_get_temp_dir(), 'tep_image');
                file_put_contents($file_name_result, file_get_contents($image_uri));

                // Заменяем изображение в архиве
                $this->zipClass->deleteName('word/' . $image_path);
                $this->zipClass->addFile($file_name_result, 'word/' . $image_path);

                // TODO: Необходимо удалить файл, но уже после того, как документ будет создан
                // unlink($file_name_result);

                return $image_xml;
            },
            $documentPart,
        ) ?? $documentPart;
    }

    /**
     * Заполняет документ данными.
     * @param  array       $data Данные
     * @throws LoaderError
     * @throws SyntaxError
     */
    public function fillWithData(array $data): void
    {
        $xml_for_twig = $this->tempDocumentMainPart;

        // Вынос Twig-тегов из XML-тегов Word
        foreach ($this->TAGS_TO_MOVE_OUT as $tag_to_move_out) {
            $xml_for_twig = $this->moveOutMacrosFromTag($xml_for_twig, $tag_to_move_out);
        }

        // Вертикально объединяем одинаковые ячейки таблицы, которые содержат "{% vm %}" (применимо для двойного цикла, где внутренний цикл объединит значения своих ячеек)
        $xml_for_twig = $this->combineTableRows($xml_for_twig);

        $xml_for_twig = $this->prepareImages($xml_for_twig);

        // Инициализация XML для Twig
        $twig_environment = new Environment(new ArrayLoader());
        $twig_template = $twig_environment->createTemplate($xml_for_twig);
        // Заполнение данными
        $xml_for_twig = $twig_template->render($data);

        $xml_for_twig = $this->insertImages($xml_for_twig);

        $this->tempDocumentMainPart = $xml_for_twig;
    }
}
