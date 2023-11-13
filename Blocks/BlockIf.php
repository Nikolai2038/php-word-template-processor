<?php

namespace App\Helpers\PhpWordHelper\Blocks;

use App\Helpers\PhpWordHelper\BlockDouble;
use Exception;

/**
 * Блок проверки if (одинарный тег).
 */
class BlockIf extends BlockDouble
{
    /**
     * Возвращает условие вида "if (<условие>)".
     * @return string
     */
    private function GetConditionCheckCommand(): string
    {
        $tag_without_id = self::GetTagWithoutId($this->tag_with_id);

        // Заменяем в строчке все специальные символы на пробелы, чтобы можно было разделить строчку
        $replaces = [
            '<',
            '>',
            '=',
            '<=',
            '>=',
            '==',
            '===',
            '!=',
            '!==',
            '<>',
        ];
        $tag_words = explode(' ', str_replace($replaces, ' ', $tag_without_id));

        // TODO: Более точный захват специальных слов и переменных
        foreach ($tag_words as $word) {
            // Замена специальных слов на PHP-код
            if ($word === 'or') {
                $new_word = ' || ';
                $tag_without_id = str_replace(' ' . $word . ' ', $new_word, $tag_without_id);
            } elseif ($word === 'and') {
                $new_word = ' && ';
                $tag_without_id = str_replace(' ' . $word . ' ', $new_word, $tag_without_id);
            } elseif ($word === 'not') {
                $new_word = ' ! ';
                $tag_without_id = str_replace(' ' . $word . ' ', $new_word, $tag_without_id);
            } // Всё остальное является переменными - преобразовываем их из строчек с точками в PHP-код
            elseif ($word !== self::RESERVED_KEYWORD_IF) {
                $new_word = self::ConvertVariableNameWithDotsToPhpCode(
                    str_replace('(', '', str_replace(')', '', $word))
                );
                $tag_without_id = str_replace(' ' . $word, ' ' . $new_word, $tag_without_id);
            }
        }

        return str_replace(self::RESERVED_KEYWORD_IF . ' ', self::RESERVED_KEYWORD_IF . ' (', $tag_without_id) . ')';
    }

    /**
     * Проверяет условие тега.
     * @param array|null $data Массив данных, из которого брать значения переменных
     * @return bool       True - условие верно, false - условие неверно
     */
    public function CheckCondition(?array $data): bool
    {
        try {
            $result = false;
            $command = $this->GetConditionCheckCommand() . ' { $result = true; } else { $result = false; }';
            eval($command);
            return $result;
        } catch (Exception $e) {
            // Перехватываем исключения, при которых используется неопределённая переменная - она делает условие равным null
            if (str_contains($e->getMessage(), 'Warning: Use of undefined constant') ||
                str_contains($e->getMessage(), 'Notice: Undefined index') ||
                // Если $data === null
                str_contains($e->getMessage(), 'Notice: Trying to access array offset on value of type null')) {
                return false;
            } else {
                throw $e;
            }
        }
    }
}
