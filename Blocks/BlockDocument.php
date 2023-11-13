<?php

namespace App\Helpers\PhpWordHelper\Blocks;

use App\Helpers\PhpWordHelper\Block;
use App\Helpers\PhpWordHelper\BlockDouble;
use App\Helpers\PhpWordHelper\PhpWordTemplateProcessor;
use Exception;

/**
 * Блок всего документа (нет тега).
 */
class BlockDocument extends BlockDouble
{
    /** Класс для работы с объектами Word */
    private PhpWordTemplateProcessor $template_processor;

    /**
     * Создаёт блок документа (единственный блок на весь документ, скрытый).
     */
    public function __construct(PhpWordTemplateProcessor $template_processor)
    {
        parent::__construct('Скрытый блок всего документа', null);

        $this->template_processor = $template_processor;

        // Первоначальная обработка переменных (убирание лишних пробелов) и их проверка
        $this->CheckAllTags();

        // Присвоение уникального ID каждой переменной в шаблоне (для двойных тегов одинаковый тег, но добавочный "/" в начале второго тега)
        $this->PrepareTagsDublicates();

        // Заменяет строки таблиц, в которых находятся двойные теги, на обычные абзацы
        // (Так как клонирование блока, начинающегося в таблице, ломает шаблон)
        $this->template_processor->ReplaceBlocksDoubleInTables();

        $variables = $this->template_processor->getVariables();
        // Создаём все объекты блоков
        $this->LoadBlocks($variables, $this->template_processor);
    }

    /**
     * Первоначальная обработка переменных и их проверка.
     */
    private function CheckAllTags(): void
    {
        // Необходимо вызывать метод каждый раз, чтобы перегенерировать переменные
        $variables = $this->template_processor->getVariables();
        $variables_count = count($variables);
        for ($i = 0; $i < $variables_count; ++$i) {
            $variable = $variables[$i];
            $variable_fixed = $variable;
            // Убирание лишних пробелов (сначала по бокам, потом внутри)
            $variable_fixed = preg_replace('/\s+/', ' ', trim($variable_fixed));
            // Убирание скобок в циклах (они лишние - данные разделяются пробелами)
            if (str_starts_with($variable, self::RESERVED_KEYWORD_FOREACH) || str_starts_with($variable, self::RESERVED_KEYWORD_FOR)) {
                $variable_fixed = str_replace('(', '', $variable_fixed);
                $variable_fixed = str_replace(')', '', $variable_fixed);
            }
            // Если тег был изменён - его обновление в шаблоне
            if ($variable !== $variable_fixed) {
                $this->template_processor->setValue($variable, '${' . $variable_fixed . '}', 1);
                $variable = $variable_fixed;
            }
        }
    }

    /**
     * Добавляет для каждого тега в шаблоне уникальный ID ("${tag}" -> "${#<id>#tag}")
     * (Необходимо для того, чтобы метод getVariables() захватывал все существующие теги, а не только уникальные).
     */
    private function PrepareTagsDublicates(): void
    {
        /** @var int $tags_count Количество проиндексированных тегов */
        $tags_count = 0;
        $opened_tags = [];
        $variables = $this->template_processor->getVariables();
        $variables_count = count($variables);
        for ($i = 0; $i < $variables_count; ++$i) {
            $variable = $variables[$i];
            // Если начинается со слеша - значит, закрывающий тег
            if (str_starts_with($variable, self::RESERVED_KEYWORD_CLOSING_SLASH)) {
                if (empty($opened_tags)) {
                    throw new Exception("Количество закрывающих тегов больше количества открывающих! Проблемная переменная: '$variable'. Последний обработанный (не окончательный ID: $i");
                } else {
                    // Устанавливаем его название точь в точь с открывающим, добавляя слеш в начало
                    $variable_new = self::RESERVED_KEYWORD_CLOSING_SLASH . array_pop($opened_tags);
                    // Изменение тега в шаблоне
                    $this->template_processor->setValue($variable, '${' . $variable_new . '}', 1);
                }
            } else {
                // Добавление ID
                $variable_new = '#' . $tags_count++ . '#' . $variable;
                // Изменение тега в шаблоне
                $this->template_processor->setValue($variable, '${' . $variable_new . '}', 1);
                // Если начинается с зарезервированного слова-блока
                if ($this->IsStartWithReserved($variable)) {
                    // Временно сохраняем открывающий тег
                    array_push($opened_tags, $variable_new);
                }
            }
            // Пересчёт массива переменных (однако индекс сохраняется, так как значения до текущего не меняются)
            $variables = $this->template_processor->getVariables();
            $variables_count = count($variables);
        }
        if (!empty($opened_tags)) {
            throw new Exception("Количество закрывающих тегов меньше количества открывающих! Проверьте теги! Не закрытые теги: '" . implode("', '", $opened_tags) . "'");
        }
    }

    /**
     * Проверяет, начинается ли название тега с зарезервированного слова двойного блока.
     * @param string $tag Переменная
     * @return bool   True - является, false - не является
     */
    private function IsStartWithReserved(string $tag): bool
    {
        foreach (Block::BLOCKS_DOUBLE_TAGS_STARTS_WITH as $reserved_tag) {
            if (str_starts_with($tag, $reserved_tag)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Обрабатывает и заполняет шаблон данными.
     * @param array|null $data Массив данных, из которого берутся данные для заполнения
     */
    public function FillData(?array $data): void
    {
        // Применение псевдонимов двойных блоков
        parent::ApplyAliases($this->template_processor);
        // Обрабатываем все внутренние блоки (генерирует и раскрывает циклы, заполняет индексы для массивов)
        parent::ProccessBlocks($this->template_processor, $data);
        // Устанавливает значения для всех переменных
        parent::ProcessGlobalValues($this->template_processor, $data);
    }
}
