<?php

namespace App\Helpers\PhpWordHelper;

use Exception;
use PhpOffice\PhpWord\TemplateProcessor;

/**
 * Класс для работы с объектами Word.
 * {@inheritdoc}
 */
class PhpWordTemplateProcessor extends TemplateProcessor
{
    // (Перегруженная функция)
    protected function replaceClonedVariables($variableReplacements, $xmlBlock)
    {
        $results = [];
        // Для каждого элемента массива сущностей
        foreach ($variableReplacements as $replacementArray) {
            $results[] = $this->replaceClonedVariablesWithArray($replacementArray, $xmlBlock);
        }
        return $results;
    }

    /**
     * Рекурсивно заменяет переменные в XML-блоке (этого функционала нет в TemplateProcessor).
     * @param array  $rereplacementArray Массив значений для замены (любой глубины, ключи будут разделяться точками)
     * @param string $xmlBlock           XML-блок
     * @param string $prefix             Префикс для рекурсивного вызова (ключи более высоких уровней массива, разделённые точками)
     */
    private function replaceClonedVariablesWithArray(
        array $replacementArray,
        string $xmlBlock,
        string $prefix = ''
    ) {
        // Для каждого поля
        foreach ($replacementArray as $search => $replacement) {
            // Если массив - вызываем рекурсивно, добавляя префикс
            if (is_array($replacement)) {
                $xmlBlock = $this->replaceClonedVariablesWithArray($replacement, $xmlBlock, $search . '.');
            }
            // Иначе - выполняем обычную замену (взято из TemplateProcessor, но добавлено использование префикса)
            else {
                $xmlBlock = $this->setValueForPart(self::ensureMacroCompleted($prefix . $search), $replacement, $xmlBlock, self::MAXIMUM_REPLACEMENTS_DEFAULT);
            }
        }
        return $xmlBlock;
    }

    /**
     * Заменяет строки таблиц, в которых находятся двойные теги, на обычные абзацы
     * (Так как клонирование блока, начинающегося в таблице, ломает шаблон).
     */
    public function ReplaceBlocksDoubleInTables(): void
    {
        $tags = $this->getVariables();
        foreach ($tags as $tag) {
            // Является ли указанный тег частью двойного блока
            if (Block::IsADoubleBlockPart($tag)) {
                /** @var bool|int[] Границы XML-блока строки таблицы, в котором лежит тег */
                $where_tag = $this->findContainingXmlBlockForMacro($tag, 'w:tr');
                // Если блок лежит в строчке таблицы
                if (is_array($where_tag)) {
                    /** @var string Содержимое абзаца, в котором находится тег (абзац лежит внутри строчки таблицы) */
                    $tag_content_in_tr = $this->getSlice($where_tag['start'], $where_tag['end']);
                    $tags_count_in_found_content = substr_count($tag_content_in_tr, '${');
                    // Если в этой строчке таблицы находится только один тег
                    if ($tags_count_in_found_content === 1) {
                        /** @var bool|int[] Границы XML-блока абзаца, в котором лежит тег */
                        $where_tag_in_p = $this->findContainingXmlBlockForMacro($tag, 'w:p');

                        /** @var string Содержимое абзаца, в котором находится тег (абзац лежит внутри строчки таблицы) */
                        $tag_content_in_p = $this->getSlice($where_tag_in_p['start'], $where_tag_in_p['end']);

                        // Заменяем строчку таблицы с тегом на абзац с тегом, чтобы при клонировании блока документ не ломался
                        $this->replaceXmlBlock($tag, $tag_content_in_p, 'w:tr');
                    }
                }
            }
        }
    }

    /**
     * Удаляет указанный блок из документа (от открывающего тега, до закрывающего).
     * @param  string $blockname Название тега
     * @return bool   True - в случае успешного удаления, false - в случае, если ничего не менялось
     */
    public function CustomRemoveBlock(string $blockname): bool
    {
        // Старый способ (думаю, немного быстрее (хотя разницы не заметил), но менее надёжнее)
        // $preg_match_condition = '/(<w:p (?:.(?!<w:p ))+\${' . $blockname . '}.*\${\/' . $blockname . '}<\/w:.*?p>)/is';
        // $replacements = 0;
        // $this->tempDocumentMainPart = preg_replace($preg_match_condition, '', $this->tempDocumentMainPart, 1, $replacements);
        // return $replacements > 0;

        /** @var bool|int[] Позиции XML-блока, содержащего открывающий тег */
        $open_block_position = $this->findContainingXmlBlockForMacro($blockname);
        if (!$open_block_position) {
            throw new Exception("Для тега '$blockname' не был найден содержащий его XML-блок");
        }

        /** @var bool|int[] Позиции XML-блока, содержащего закрывающий тег */
        $close_block_position = $this->findContainingXmlBlockForMacro(Block::RESERVED_KEYWORD_CLOSING_SLASH . $blockname);
        if (!$close_block_position) {
            throw new Exception("Для тега '" . Block::RESERVED_KEYWORD_CLOSING_SLASH . "$blockname' не был найден содержащий его XML-блок");
        }

        // Находим содержимое между блоками, включая сами блоки
        $xml_content = $this->getSlice($open_block_position['start'], $close_block_position['end']);

        $replacements = 0;
        $this->tempDocumentMainPart = str_replace($xml_content, '', $this->tempDocumentMainPart, $replacements);
        // Дополнительная проверка
        if ($replacements > 1) {
            throw new Exception("При удалении блока '$blockname' замен было больше, чем должно было быть! Что-то пошло не так");
        }

        return $replacements === 1;
    }

    /**
     * Дублирует указанный блок указанное количество раз.
     * @param  string $blockname    Название тега блока
     * @param  int    $blocks_count Количество блоков в результате
     * @return bool   True - в случае успешного клонирования, false - в случае, если ничего не менялось
     */
    public function CustomCloneBlock(string $blockname, int $blocks_count): bool
    {
        // Если клонирований нет - просто удаляем блок
        if ($blocks_count < 1) {
            return $this->CustomRemoveBlock($blockname);
        }

        /** @var bool|int[] Позиции XML-блока, содержащего открывающий тег */
        $open_block_position = $this->findContainingXmlBlockForMacro($blockname);
        if (!$open_block_position) {
            throw new Exception("Для тега '$blockname' не был найден содержащий его XML-блок");
        }
        // Получаем содержимое XML-блока, содержащего открывающий тег, чтобы потом его удалить
        // (Важно получить содержимое до того, как шаблон поменяется и позиции сменятся)
        $xml_open_block = $this->getSlice($open_block_position['start'], $open_block_position['end']);

        /** @var bool|int[] Позиции XML-блока, содержащего закрывающий тег */
        $close_block_position = $this->findContainingXmlBlockForMacro(Block::RESERVED_KEYWORD_CLOSING_SLASH . $blockname);
        if (!$close_block_position) {
            throw new Exception("Для тега '" . Block::RESERVED_KEYWORD_CLOSING_SLASH . "$blockname' не был найден содержащий его XML-блок");
        }
        // Получаем содержимое XML-блока, содержащего закрывающий тег, чтобы потом его удалить
        // (Важно получить содержимое до того, как шаблон поменяется и позиции сменятся)
        $xml_close_block = $this->getSlice($close_block_position['start'], $close_block_position['end']);

        // Находим содержимое внутри блока и клонируем его нужное число раз
        $xml_inside_block = $this->getSlice($open_block_position['end'], $close_block_position['start']);
        $new_xml_inside_block = str_repeat($xml_inside_block, $blocks_count);

        // Сама замена
        $replacements_all = 0;
        $replacements = 0;
        $this->tempDocumentMainPart = str_replace($xml_open_block, '', $this->tempDocumentMainPart, $replacements);
        $replacements_all += $replacements;
        $this->tempDocumentMainPart = self::StrReplaceFirst($xml_inside_block, $new_xml_inside_block, $this->tempDocumentMainPart, $replacements);
        $replacements_all += $replacements;
        $this->tempDocumentMainPart = str_replace($xml_close_block, '', $this->tempDocumentMainPart, $replacements);
        $replacements_all += $replacements;

        // Дополнительная проверка
        if ($replacements_all > 3) {
            throw new Exception("При клонировании блока '$blockname' замен было больше, чем должно было быть! Что-то пошло не так");
        }
        return $replacements_all === 3;
    }

    /**
     * Заменяет подстрочку в строчке лишь один раз.
     * @param  string $needle       Подстрочка
     * @param  string $replace      На что заменить
     * @param  string $haystack     Исходный текст
     * @param  int    $replacements Количество замен (0, если не было, 1 - если были)
     * @return string Изменённая строчка
     */
    private static function StrReplaceFirst(string $needle, string $replace, string $haystack, int &$replacements = 0): string
    {
        $pos = strpos($haystack, $needle);
        if ($pos !== false) {
            $replacements = 1;
            return substr_replace($haystack, $replace, $pos, strlen($needle));
        } else {
            $replacements = 0;
            return $haystack;
        }
    }
}
