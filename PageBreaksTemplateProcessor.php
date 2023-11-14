<?php

namespace App\Helpers\PhpWord;

use PhpOffice\PhpWord\TemplateProcessor;

/**
 * Класс для работы с объектами Word.
 * @see TemplateProcessor
 */
class PageBreaksTemplateProcessor extends TemplateProcessor
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
}
