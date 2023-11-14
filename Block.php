<?php

namespace App\Helpers\PhpWordHelper;

use App\Helpers\PhpWord\Blocks\BlockForeach;

/**
 * Общий родительский класс для всех блоков.
 */
abstract class Block
{
    /** @var string Ключевое слово, с которого начинается тег цикла foreach (один вариант) */
    protected const RESERVED_KEYWORD_FOREACH = 'foreach';
    /** @var string Ключевое слово, с которого начинается тег цикла foreach (второй вариант) */
    protected const RESERVED_KEYWORD_FOR = 'for';
    /** @var string Ключевое слово, с которого начинается тег условия if */
    protected const RESERVED_KEYWORD_IF = 'if';
    /** @var string Ключевое слово, с которого начинается тег image (переменная, но контент становится изображением) */
    protected const RESERVED_KEYWORD_IMAGE = 'image';
    /** @var string Ключевое слово, с которо начинается закрывающий тег любого двойного блока (цикл foreach, условие if и пр.) */
    public const RESERVED_KEYWORD_CLOSING_SLASH = '/';
    /** @var string Ключевое слово, разделяющее название массива и псевдоним переменной в теге цикла foreach (первый способ объявления цикла) */
    protected const RESERVED_KEYWORD_AS = 'as';
    /** @var string Ключевое слово, разделяющее псевдоним переменной и название массива в теге цикла foreach (второй способ объявления цикла) */
    protected const RESERVED_KEYWORD_IN = 'in';

    /** @var string[] Зарезервированные ключевые слова, с которых начинаются двойные блоки */
    public const BLOCKS_DOUBLE_TAGS_STARTS_WITH = [
        self::RESERVED_KEYWORD_FOREACH,
        self::RESERVED_KEYWORD_FOR,
        self::RESERVED_KEYWORD_IF,
    ];

    /** @var string[] Зарезервированные ключевые слова, которые нельзя использовать как ключи для данных */
    public const RESERVED_KEYWORDS = [
        self::RESERVED_KEYWORD_FOREACH,
        self::RESERVED_KEYWORD_FOR,
        self::RESERVED_KEYWORD_IF,
        self::RESERVED_KEYWORD_IMAGE,
        self::RESERVED_KEYWORD_CLOSING_SLASH,
        self::RESERVED_KEYWORD_AS,
    ];

    /** @var string Название тега вместе с ID (в формате "#<ID>#<Название>") */
    protected string $tag_with_id;

    /**
     * Создаёт блок.
     * @param string $tag_with_id Название тега
     */
    protected function __construct(string $tag_with_id)
    {
        $this->tag_with_id = $tag_with_id;
    }

    /**
     * Возвращает название тега без указания ID.
     * @param string $tag_with_or_without_id Название тега с указанным ("#<ID>#<Название>") или не указанным ("<Название>") ID
     * @return string Название тега без ID ("<Название>")
     */
    public static function GetTagWithoutId(string $tag_with_or_without_id): string
    {
        if (str_starts_with($tag_with_or_without_id, '#')) {
            return explode('#', $tag_with_or_without_id, 3)[2];
        } else {
            return $tag_with_or_without_id;
        }
    }

    /**
     * Проверяет является ли указанный тег указанным типом блока.
     * @param string $tag Название тега с указанным ("#<ID>#<Название>") или не указанным ("<Название>") ID
     * @param string $block_name Зарезервированное название блока
     * @return bool   True - является, false - нет
     */
    public static function IsA(string $tag, string $block_name): bool
    {
        return str_starts_with(self::GetTagWithoutId($tag), $block_name);
    }

    /**
     * Проверяет является ли указанный тег частью двойного блока.
     * @param string $tag Название тега с указанным ID ("#<ID>#<Название>") или не указанным ID ("<Название>")
     * @return bool   True - является, false - нет
     */
    public static function IsADoubleBlockPart(string $tag): bool
    {
        foreach (self::BLOCKS_DOUBLE_TAGS_STARTS_WITH as $keyword) {
            if (self::IsA($tag, $keyword)) {
                return true;
            }
        }
        if (self::IsA($tag, self::RESERVED_KEYWORD_CLOSING_SLASH)) {
            return true;
        }
        return false;
    }

    /**
     * Конвертирует ключ переменной, использующий разделители в виде точек, в прямое обращение к элементу массива PHP.
     * @param string $variable_full_name_as_dots Ключ переменной в файле шаблона
     * @return string Обращение к переменной в массиве, как это бы выглядело в коде PHP
     */
    protected static function ConvertVariableNameWithDotsToPhpCode(string $variable_full_name_as_dots): string
    {
        // Замена "users[0].name" на "users.0.name"
        $variable_full_name_as_dots = str_replace('[', '.', str_replace(']', '', $variable_full_name_as_dots));
        /** @var string[] Ключи переменной (которые были разделены точками) */
        $variable_keys = explode('.', $variable_full_name_as_dots);

        // Генерация названия PHP-переменной для поиска среди данных
        /** @var string Название перменной в PHP-коде */
        $variable_full_name_as_php = '$data';
        foreach ($variable_keys as $key) {
            $variable_full_name_as_php .= "['$key']";
        }

        return $variable_full_name_as_php;
    }

    /**
     * Возвращает значение PHP-переменной по её названию в шаблоне Word.
     * @param string $variable_full_name_as_dots Ключ переменной в файле шаблона
     * @param array|null $data Массив данных, в котором искать переменную
     * @return array|string|null Значение переменной
     */
    protected static function GetVariableValue(
        string $variable_full_name_as_dots,
        // Пусть она и не подсвечивается, эта переменная используется в eval ниже!
        ?array $data
    )
    {
        $variable_full_name_as_php = self::ConvertVariableNameWithDotsToPhpCode($variable_full_name_as_dots);

        // Вызов команды для генерации подставляемого значения
        $command = "\$value = isset($variable_full_name_as_php) ? $variable_full_name_as_php : null;";
        eval($command);

        return $value ?? null;
    }

    /**
     * Заменяет все теги в указанном блоке, а также в дочерних блоках, заменяя индекс для массива.
     * @param CustomTemplateProcessor $template_processor Класс для работы с объектами Word
     * @param string $foreach_array_name Название массива
     * @param string $foreach_variable_name Название псевдонима переменной массива
     * @param int $index Индекс, который необходимо встроить
     * @return mixed
     */
    protected function ReplaceForeachIndex(
        CustomTemplateProcessor $template_processor,
        string                  $foreach_array_name,
        string                  $foreach_variable_name,
        int                     $index
    ): void
    {
        /** @var string Старое название тега */
        $old_tag = $this->tag_with_id;
        /** @var string Новое название тега */
        $new_tag = $this->GetTagWithReplacedForeachIndex($this->tag_with_id, $foreach_array_name, $foreach_variable_name, $index);
        // Замена тега в шаблоне
        $template_processor->setValue(
            $this->tag_with_id,
            '${' . $new_tag . '}',
            1
        );
        // Установка тега в свойствах объекта
        $this->tag_with_id = $new_tag;

        // Если блок ещё и двойной
        if (is_a($this, BlockDouble::class)) {
            /** @var BlockDouble $this */
            // Замена и закрывающего тега в шаблоне
            $template_processor->setValue(
                self::RESERVED_KEYWORD_CLOSING_SLASH . $old_tag,
                '${' . self::RESERVED_KEYWORD_CLOSING_SLASH . $new_tag . '}',
                1
            );

            // Если массив - необходимо заменить и свойства "foreach_array_name" и "foreach_variable_name"
            if (is_a($this, BlockForeach::class)) {
                /** @var BlockForeach $this */
                /** @var string Новое название массива */
                $new_foreach_array_name = $this->GetTagWithReplacedForeachIndex($this->foreach_array_name, $foreach_array_name, $foreach_variable_name, $index);
                $this->foreach_array_name = $new_foreach_array_name;

                /** @var string Новое название псевдонима массива */
                $new_foreach_variable_name = $this->GetTagWithReplacedForeachIndex($this->foreach_variable_name, $foreach_array_name, $foreach_variable_name, $index);
                $this->foreach_variable_name = $new_foreach_variable_name;
            }

            // Для всех двойных блоков так же проходим и по дочерним блокам
            foreach ($this->blocks as $block) {
                $block->ReplaceForeachIndex($template_processor, $foreach_array_name, $foreach_variable_name, $index);
            }
        }
    }

    /**
     * Возвращает название тега, с заменённым индексом для массива.
     * @param string $tag Исходное название тега
     * @param string $foreach_array_name Название массива
     * @param string $foreach_variable_name Название псевдонима переменной массива
     * @param int $index Индекс, который необходимо встроить
     * @return string Новое название тега
     */
    protected function GetTagWithReplacedForeachIndex(
        string $tag,
        string $foreach_array_name,
        string $foreach_variable_name,
        int    $index
    ): string
    {
        $new_tag = $tag;
        // Заменяем псевдоним переменной массива ("user." -> "users[<id>].")
        $new_tag = str_replace(
            $foreach_variable_name . '.',
            $foreach_array_name . '[' . $index . '].',
            $new_tag
        );
        // Заменяем псевдоним переменной массива ("user " -> "users[<id>]")
        $new_tag = str_replace(
            ' ' . $foreach_variable_name . ' ',
            ' ' . $foreach_array_name . '[' . $index . '] ',
            $new_tag
        );
        // TODO: Разобраться, почему не работает
        // Обработка псевдонима в конце тега (случай не точки и не пробела)
        // if (str_ends_with($new_tag, $foreach_variable_name)) {
        //     $new_tag = str_replace(
        //         $foreach_variable_name,
        //         $foreach_array_name . '[' . $index . ']',
        //         $new_tag
        //     );
        // }
        return $new_tag;
    }

    /**
     * Возвращает блок-копию данного (новый объект).
     * @return Block
     */
    protected function GetCopy(): Block
    {
        return clone $this;
    }

    /**
     * Возвращает копию массива блоков (новые объекты).
     * @param array<Block> $blocks Исходный массив блоков
     * @return array<Block> Скопированный массив блоков (новые объекты)
     */
    protected static function GetCopyOfArray(array $blocks)
    {
        $result = [];
        foreach ($blocks as $block) {
            $result[] = $block->GetCopy();
        }
        return $result;
    }
}
