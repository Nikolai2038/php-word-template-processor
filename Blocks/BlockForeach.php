<?php

namespace App\Helpers\PhpWordHelper\Blocks;

use App\Helpers\PhpWordHelper\BlockDouble;
use Exception;

/**
 * Блок цикла foreach (двойной тег).
 */
class BlockForeach extends BlockDouble
{
    /** @var string Название массива, из которого берутся данные */
    public string $foreach_array_name;

    /** @var string Псевдоним переменной массива, из которого берутся данные */
    public string $foreach_variable_name;

    /**
     * Создаёт блок foreach.
     * @param string $tag_with_id Название тега
     */
    protected function __construct(string $tag_with_id, BlockDouble $parent)
    {
        parent::__construct($tag_with_id, $parent);

        $tag_without_id = self::GetTagWithoutId($tag_with_id);
        // Убираем из тега скобки и разбиваем на части: "foreach <массив> as <псевдоним>"
        $tag_data = explode(' ', str_replace(')', '', str_replace('(', '', $tag_without_id)));
        if (count($tag_data) !== 4 || ($tag_data[0] !== self::RESERVED_KEYWORD_FOREACH && $tag_data[0] !== self::RESERVED_KEYWORD_FOR)) {
            throw new Exception("Синтаксическая ошибка в шаблоне Word! Неправильный синтаксис foreach в теге '$tag_without_id'");
        }

        // Если условие цикла разделяется ключевым словом "as", то сначала идёт название массива, а потом название переменной
        if ($tag_data[2] === self::RESERVED_KEYWORD_AS) {
            $this->foreach_array_name = $tag_data[1];
            $this->foreach_variable_name = $tag_data[3];
        } // Если условие цикла разделяется ключевым словом "in", то сначала идёт название переменной, а потом название массива
        elseif ($tag_data[2] === self::RESERVED_KEYWORD_IN) {
            $this->foreach_array_name = $tag_data[3];
            $this->foreach_variable_name = $tag_data[1];
        } // Иначе - синтаксическая ошибка
        else {
            throw new Exception("Синтаксическая ошибка в шаблоне Word! Неправильный синтаксис foreach в теге '$tag_without_id'. Неизвестное слово-разделитель '$tag_data[2]'");
        }

        if (array_key_exists($this->foreach_variable_name, $this->aliases)) {
            throw new Exception("Синтаксическая ошибка в шаблоне Word! Используется уже существующий псевдоним '" . $this->foreach_variable_name . "' в теге '$tag_without_id'");
        }
    }
}
