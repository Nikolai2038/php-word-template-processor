<?php

namespace App\Helpers\PhpWordHelper;

use Exception;
use PhpOffice\PhpWord\TemplateProcessor;

/**
 * Класс для работы с объектами Word.
 * @see TemplateProcessor
 */
class CustomTemplateProcessor extends TemplateProcessor
{
    /** XML для разрыва страницы */
    private const PAGE_BREAK_XML = '<w:br w:type="page"/>';

    /** Максимальное количество разрывов страниц (предохранение от бесконечного цикла) */
    private const LIMIT_WHILE_ITERATIONS = 1000;

    /**
     * Заменяет все значения указанной переменной на разрывы страниц.
     * При этом, если такая переменная расположена в таблице, таблица будет разделена
     * (строчка, в которой была расположена переменная разрыва страницы, будет помещена на новую страницу).
     * @param string $variable_name Название переменной, которая будет заменяться на разрыв страницы
     */
    public function ReplacePageBreakVariable(string $variable_name): void
    {
        $while_iteration_id = 0;

        // Пока в тексте есть переменные переноса строк
        while ($this->getVariableWithNameCount($variable_name) > 0 && $while_iteration_id++ < self::LIMIT_WHILE_ITERATIONS) {
            // Получаем границы XML-блока строки таблицы, в котором лежит переменная
            $where_tag_in_tr = $this->findContainingXmlBlockForMacro($variable_name, 'w:tr');

            // Если блок не лежит в строчке таблицы
            if (!is_array($where_tag_in_tr)) {
                // Просто заменяем на разрыв страницы и продолжаем цикл
                $this->replaceXmlBlock($variable_name, self::PAGE_BREAK_XML, 'w:t');
                continue;
            }

            // ========================================
            // Получение описания (заголовка) таблицы
            // ========================================
            // Получаем границы XML-блока таблицы, в котором лежит переменная
            $where_tag_in_tbl = $this->findContainingXmlBlockForMacro($variable_name, 'w:tbl');

            // Получаем содержимое таблицы
            $tag_content_in_tbl = $this->getSlice($where_tag_in_tbl['start'], $where_tag_in_tbl['end']);

            $matches = [];
            preg_match_all('/(<w:tblPr>.*<\/w:tblGrid>)/i', $tag_content_in_tbl, $matches);
            $table_header = $matches[1][0];
            // ========================================

            // Получаем содержимое строчки таблицы
            $tag_content_in_tr = $this->getSlice($where_tag_in_tr['start'], $where_tag_in_tr['end']);

            // Из строчки таблицы убираем переменную
            $tag_content_in_tr = str_replace('<w:t>${' . $variable_name . '}</w:t>', '', $tag_content_in_tr);

            // Добавляем абзац с разрывом страницы перед строчкой таблицы, в котором указана переменная
            $this->replaceXmlBlockRegex($variable_name, '</w:tbl><w:p><w:r>' . self::PAGE_BREAK_XML . '</w:r></w:p><w:tbl>' . $table_header . $tag_content_in_tr, 'w:tr');
        }

        if ($while_iteration_id >= self::LIMIT_WHILE_ITERATIONS) {
            throw new RuntimeException("Цикл while был прерван, так как выполнился уже " . $while_iteration_id . " раз!");
        }
    }

    /**
     * Заменяет все значения переменных, удоавлетворяющих указанному регулярному выражению, на разрывы страниц.
     * При этом, если такая переменная расположена в таблице, таблица будет разделена
     * (строчка, в которой была расположена переменная разрыва страницы, будет помещена на новую страницу).
     */
    public function ReplacePageBreakVariableRegex(string $variable_name_regex): void
    {
        $while_iteration_id = 0;

        // Пока в тексте есть переменные переноса строк
        while ($this->getVariableWithNameCount($variable_name_regex) > 0 && $while_iteration_id++ < self::LIMIT_WHILE_ITERATIONS) {
            // Получаем границы XML-блока строки таблицы, в котором лежит переменная
            $where_tag_in_tr = $this->findContainingXmlBlockForMacroRegex($variable_name_regex, 'w:tr');

            // Если блок не лежит в строчке таблицы
            if (!is_array($where_tag_in_tr)) {
                // Просто заменяем на разрыв страницы и продолжаем цикл
                $this->replaceXmlBlockRegex($variable_name_regex, self::PAGE_BREAK_XML, 'w:t');
                continue;
            }

            // ========================================
            // Получение описания (заголовка) таблицы
            // ========================================
            // Получаем границы XML-блока таблицы, в котором лежит переменная
            $where_tag_in_tbl = $this->findContainingXmlBlockForMacroRegex($variable_name_regex, 'w:tbl');

            // Получаем содержимое таблицы
            $tag_content_in_tbl = $this->getSlice($where_tag_in_tbl['start'], $where_tag_in_tbl['end']);

            $matches = [];
            preg_match_all('/(<w:tblPr>.*<\/w:tblGrid>)/i', $tag_content_in_tbl, $matches);
            $table_header = $matches[1][0];
            // ========================================

            // Получаем содержимое строчки таблицы
            $tag_content_in_tr = $this->getSlice($where_tag_in_tr['start'], $where_tag_in_tr['end']);

            // Из строчки таблицы убираем переменную
            $tag_content_in_tr = preg_replace('/<w:t>\${' . $variable_name_regex . '}<\/w:t>/', '', $tag_content_in_tr);

            // Добавляем абзац с разрывом страницы перед строчкой таблицы, в котором указана переменная
            $this->replaceXmlBlockRegex($variable_name_regex, '</w:tbl><w:p><w:r>' . self::PAGE_BREAK_XML . '</w:r></w:p><w:tbl>' . $table_header . $tag_content_in_tr, 'w:tr');
        }

        if ($while_iteration_id >= self::LIMIT_WHILE_ITERATIONS) {
            throw new RuntimeException("Цикл while был прерван, так как выполнился уже " . $while_iteration_id . " раз!");
        }
    }

    /**
     * (Своя реализация метода с указанием конкретного названия переменной для улучшения производительности).
     * @see TemplateProcessor::getVariableCount
     */
    public function getVariableWithNameCount(string $variableName): int
    {
        $variables = $this->getVariablesWithNameForPart($this->tempDocumentMainPart, $variableName);

        foreach ($this->tempDocumentHeaders as $headerXML) {
            $variables = array_merge(
                $variables,
                $this->getVariablesWithNameForPart($headerXML, $variableName),
            );
        }

        foreach ($this->tempDocumentFooters as $footerXML) {
            $variables = array_merge(
                $variables,
                $this->getVariablesWithNameForPart($footerXML, $variableName),
            );
        }

        return sizeof($variables);
    }

    /**
     * (Своя реализация метода с указанием конкретного названия переменной для улучшения производительности).
     * @see TemplateProcessor::getVariablesForPart
     */
    protected function getVariablesWithNameForPart(string $documentPartXML, string $variableName): array
    {
        $matches = [];
        preg_match_all('/\$\{(' . $variableName . ')}/i', $documentPartXML, $matches);
        return $matches[1];
    }

    /**
     * (Своя реализация, чтобы использовать регулярное выражение вместо конкретного названия переменной).
     * @see TemplateProcessor::findContainingXmlBlockForMacro
     */
    protected function findContainingXmlBlockForMacroRegex($macro, $blockType = 'w:p')
    {
        $macroPos = $this->findMacroRegex($macro);
        if (0 > $macroPos) {
            return false;
        }
        $start = $this->findXmlBlockStartRegex($macroPos, $blockType);
        if (0 > $start) {
            return false;
        }
        $end = $this->findXmlBlockEndRegex($start, $blockType);
        // if not found or if resulting string does not contain the macro we are searching for
        if (0 > $end || preg_match('/' . $macro . '/', $this->getSlice($start, $end)) === 0) {
            return false;
        }

        return ['start' => $start, 'end' => $end];
    }

    /**
     * (Своя реализация, чтобы использовать регулярное выражение вместо конкретного названия переменной).
     * @see TemplateProcessor::findMacro
     */
    protected function findMacroRegex($search, $offset = 0): int
    {
        preg_match('/\$\{(' . $search . ')}/i', $this->tempDocumentMainPart, $matches, PREG_OFFSET_CAPTURE, $offset);

        return empty($matches) ? -1 : $matches[0][1];
    }

    /**
     * (Своя реализация, чтобы использовать регулярное выражение вместо конкретного названия переменной).
     * @see TemplateProcessor::findXmlBlockStart
     */
    protected function findXmlBlockStartRegex($offset, $blockType): int
    {
        $reverseOffset = (strlen($this->tempDocumentMainPart) - $offset) * -1;
        // first try XML tag with attributes
        $blockStart = strrpos($this->tempDocumentMainPart, '<' . $blockType . ' ', $reverseOffset);
        // if not found, or if found but contains the XML tag without attribute
        if (false === $blockStart || strrpos($this->getSlice($blockStart, $offset), '<' . $blockType . '>')) {
            // also try XML tag without attributes
            $blockStart = strrpos($this->tempDocumentMainPart, '<' . $blockType . '>', $reverseOffset);
        }

        return ($blockStart === false) ? -1 : $blockStart;
    }

    /**
     * (Своя реализация, чтобы использовать регулярное выражение вместо конкретного названия переменной).
     * @see TemplateProcessor::findXmlBlockEnd
     */
    protected function findXmlBlockEndRegex($offset, $blockType): int
    {
        $blockEndStart = strpos($this->tempDocumentMainPart, '</' . $blockType . '>', $offset);
        // return position of end of tag if found, otherwise -1

        return ($blockEndStart === false) ? -1 : $blockEndStart + 3 + strlen($blockType);
    }

    /**
     * (Своя реализация, чтобы использовать регулярное выражение вместо конкретного названия переменной).
     * @see TemplateProcessor::replaceXmlBlock
     */
    protected function replaceXmlBlockRegex($macro, $block, $blockType = 'w:p'): self
    {
        $where = $this->findContainingXmlBlockForMacroRegex($macro, $blockType);
        if (is_array($where)) {
            $this->tempDocumentMainPart = $this->getSlice(0, $where['start']) . $block . $this->getSlice($where['end']);
        }

        return $this;
    }

    // (Перегруженная функция)
    protected function replaceClonedVariables($variableReplacements, $xmlBlock): array
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
     * @param string $xmlBlock XML-блок
     * @param string $prefix   Префикс для рекурсивного вызова (ключи более высоких уровней массива, разделённые точками)
     */
    private function replaceClonedVariablesWithArray(
        array $replacementArray,
        string $xmlBlock,
        string $prefix = '',
    ) {
        // Для каждого поля
        foreach ($replacementArray as $search => $replacement) {
            // Если массив - вызываем рекурсивно, добавляя префикс
            if (is_array($replacement)) {
                $xmlBlock = $this->replaceClonedVariablesWithArray($replacement, $xmlBlock, $search . '.');
            } // Иначе - выполняем обычную замену (взято из TemplateProcessor, но добавлено использование префикса)
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
                /** @var bool|int[] $tag Границы XML-блока строки таблицы, в котором лежит тег */
                $where_tag = $this->findContainingXmlBlockForMacro($tag, 'w:tr');
                // Если блок лежит в строчке таблицы
                if (is_array($where_tag)) {
                    /** @var string $tag_content_in_tr Содержимое абзаца, в котором находится тег (абзац лежит внутри строчки таблицы) */
                    $tag_content_in_tr = $this->getSlice($where_tag['start'], $where_tag['end']);
                    $tags_count_in_found_content = substr_count($tag_content_in_tr, '${');
                    // Если в этой строчке таблицы находится только один тег
                    if ($tags_count_in_found_content === 1) {
                        /** @var bool|int[] $tag Границы XML-блока абзаца, в котором лежит тег */
                        $where_tag_in_p = $this->findContainingXmlBlockForMacro($tag, 'w:p');

                        /** @var string $tag_content_in_p Содержимое абзаца, в котором находится тег (абзац лежит внутри строчки таблицы) */
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
     * @param  string    $blockname Название тега
     * @return bool      True - в случае успешного удаления, false - в случае, если ничего не менялось
     * @throws Exception
     */
    public function CustomRemoveBlock(string $blockname): bool
    {
        // Старый способ (думаю, немного быстрее (хотя разницы не заметил), но менее надёжнее)
        // $preg_match_condition = '/(<w:p (?:.(?!<w:p ))+\${' . $blockname . '}.*\${\/' . $blockname . '}<\/w:.*?p>)/is';
        // $replacements = 0;
        // $this->tempDocumentMainPart = preg_replace($preg_match_condition, '', $this->tempDocumentMainPart, 1, $replacements);
        // return $replacements > 0;

        /** @var bool|int[] $blockname Позиции XML-блока, содержащего открывающий тег */
        $open_block_position = $this->findContainingXmlBlockForMacro($blockname);
        if (!$open_block_position) {
            throw new Exception("Для тега '$blockname' не был найден содержащий его XML-блок");
        }

        /** @var bool|int[] $blockname Позиции XML-блока, содержащего закрывающий тег */
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
     * @param  string    $blockname    Название тега блока
     * @param  int       $blocks_count Количество блоков в результате
     * @return bool      True - в случае успешного клонирования, false - в случае, если ничего не менялось
     * @throws Exception
     */
    public function CustomCloneBlock(string $blockname, int $blocks_count): bool
    {
        // Если клонирований нет - просто удаляем блок
        if ($blocks_count < 1) {
            return $this->CustomRemoveBlock($blockname);
        }

        /** @var bool|int[] $blockname Позиции XML-блока, содержащего открывающий тег */
        $open_block_position = $this->findContainingXmlBlockForMacro($blockname);
        if (!$open_block_position) {
            throw new Exception("Для тега '$blockname' не был найден содержащий его XML-блок");
        }
        // Получаем содержимое XML-блока, содержащего открывающий тег, чтобы потом его удалить
        // (Важно получить содержимое до того, как шаблон поменяется и позиции сменятся)
        $xml_open_block = $this->getSlice($open_block_position['start'], $open_block_position['end']);

        /** @var bool|int[] $blockname Позиции XML-блока, содержащего закрывающий тег */
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
