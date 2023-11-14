<?php

namespace App\Helpers\PhpWordHelper\Blocks;

use App\Helpers\PhpWordHelper\Block;
use App\Helpers\PhpWordHelper\CustomTemplateProcessor;

/**
 * Блок переменной (одинарный тег).
 */
class BlockVariable extends Block
{
    /** @var string Разделитель между ключом переменной и значением по умолчанию */
    private const DEFAULT_VALUE_SEPARATOR = ':';

    /** @var bool Использует ли переменная псевдоним foreach */
    public bool $is_using_foreach_alias = false;

    /** @var string Значение по умолчанию, которое выставляется переменной, если её значения не было во входных данных */
    private string $default_value = '';

    /**
     * Создаёт блок переменной.
     * @param string $tag_with_id Название тега
     * @param CustomTemplateProcessor $template_processor Класс для работы с объектами Word
     */
    public function __construct(string $tag_with_id, CustomTemplateProcessor $template_processor)
    {
        $this->tag_with_id = $tag_with_id;
        // Если для переменной указано значение по умолчанию
        if (str_contains($tag_with_id, self::DEFAULT_VALUE_SEPARATOR)) {
            $tag_data = explode(self::DEFAULT_VALUE_SEPARATOR, $tag_with_id, 2);
            $this->tag_with_id = $tag_data[0];
            $this->default_value = $tag_data[1];
            // Заменяем сам тег в шаблоне, убирая значение по умолчанию (его сохранили в свойство объекта)
            $template_processor->setValue($tag_with_id, '${' . $this->tag_with_id . '}', 1);
        }
    }

    /**
     * Возвращает значение по умолчанию, которое выставляется переменной, если её значения не было во входных данных.
     * @return string
     */
    public function GetDefaultValue(): string
    {
        return $this->default_value;
    }
}
