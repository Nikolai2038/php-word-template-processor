<?php

namespace App\Helpers\PhpWordHelper;

use App\Helpers\PhpWord\Blocks\BlockForeach;
use App\Helpers\PhpWord\Blocks\BlockIf;
use App\Helpers\PhpWord\Blocks\BlockVariable;
use App\Helpers\PhpWord\Blocks\BlockVariableImage;
use Exception;

/**
 * Общий родительский класс для всех двойных блоков (имеют закрывающий тег).
 */
abstract class BlockDouble extends Block
{
    /** @var array<Block> Массив блоков внутри этого блока */
    protected array $blocks = [];

    /** @var array Ассоциативный массив псевдонимов внутри этого блока ("псевдоним" => "значение"). Идут в порядке от самих глубоких до верхних. */
    protected array $aliases = [];

    /**
     * Создаёт двойной блок.
     * @param string           $tag_with_id Название тега
     * @param BlockDouble|null $parent      Родительский блок
     */
    protected function __construct(string $tag_with_id, ?BlockDouble $parent)
    {
        parent::__construct($tag_with_id);
        if (isset($parent)) {
            $this->aliases = $parent->aliases;
        }
        // Для генерации дополнительного номера блоку при его замене
        srand(0);
    }

    /**
     * Создаёт блоки внутри этого блока на основе переданного массива тегов шаблона документа.
     * @param array                    &$tags_with_ids     Массив тегов с указанными ID в начале (для уникальности - ни один тег не будет пропущен)
     * @param PhpWordTemplateProcessor $template_processor Класс для работы с объектами Word
     */
    protected function LoadBlocks(array &$tags_with_ids, PhpWordTemplateProcessor $template_processor): void
    {
        $this->blocks = [];
        $tag_with_id = array_shift($tags_with_ids);
        // Пока теги не кончатся
        while (isset($tag_with_id)) {
            if (self::IsA($tag_with_id, self::RESERVED_KEYWORD_FOREACH) || self::IsA($tag_with_id, self::RESERVED_KEYWORD_FOR)) {
                $block = new BlockForeach($tag_with_id, $this);
            } elseif (self::IsA($tag_with_id, self::RESERVED_KEYWORD_IF)) {
                $block = new BlockIf($tag_with_id, $this);
            }
            // Если закрывающий тег - выходим из рекурсии
            elseif (self::IsA($tag_with_id, self::RESERVED_KEYWORD_CLOSING_SLASH)) {
                return;
            }
            // Если изображение
            elseif (self::IsA($tag_with_id, self::RESERVED_KEYWORD_IMAGE)) {
                $block = new BlockVariableImage($tag_with_id, $template_processor);
            }
            // В остальных случаях - обычная переменная
            else {
                $block = new BlockVariable($tag_with_id, $template_processor);
            }
            // Если блок является двойным - сначала загружаем все блоки внутри него
            if (is_a($block, BlockDouble::class)) {
                $block->LoadBlocks($tags_with_ids, $template_processor);
            }
            // Потом уже добавляем тег родительскому
            $this->blocks[] = $block;
            // Считываем следующее значение
            $tag_with_id = array_shift($tags_with_ids);
        }
    }

    /**
     * Рекурсивно применяет псевдонимы для всех тегов внутри этого блока.
     * @param PhpWordTemplateProcessor $template_processor Класс для работы с объектами Word
     */
    protected function ApplyAliases(PhpWordTemplateProcessor $template_processor): void
    {
        // Преобразование всех переменных, добавляя "[#]" для массивов
        foreach ($this->blocks as $block) {
            // Если двойной блок - необходимо сделать закрывающий тег аналогично открывающему, для которого могли применится псевдонимы
            if (is_a($block, self::class)) {
                /** @var self $block */
                // Если цикл - ему необходимо применить псевдонимы и на foreach_array_name
                if (is_a($block, BlockForeach::class)) {
                    /** @var BlockForeach $block */
                    $new_foreach_array_name = $block->foreach_array_name;
                    $new_tag_with_id = $block->tag_with_id;
                    // Применение псевдонимов
                    foreach ($this->aliases as $alias_key => $alias_value) {
                        $new_foreach_array_name = str_replace($alias_key, $alias_value, $new_foreach_array_name);
                        $new_tag_with_id = str_replace($alias_key, $alias_value, $new_tag_with_id);
                    }
                    // Замена и открывающего, и закрывающего тега в шаблоне
                    $template_processor->setValue($block->tag_with_id, '${' . $new_tag_with_id . '}', 1);
                    $template_processor->setValue(self::RESERVED_KEYWORD_CLOSING_SLASH . $block->tag_with_id, '${' . self::RESERVED_KEYWORD_CLOSING_SLASH . $new_tag_with_id . '}', 1);
                    // Обновление свойств объекта
                    $block->foreach_array_name = $new_foreach_array_name;
                    $block->tag_with_id = $new_tag_with_id;
                }
                // У двойных блоков не циклов нет (пока что) каких-либо полей, для которых нужно применить псевдоним
                else {
                    $new_tag_with_id = $block->tag_with_id;
                    // Применение псевдонимов
                    foreach ($this->aliases as $alias_key => $alias_value) {
                        $new_tag_with_id = str_replace($alias_key, $alias_value, $new_tag_with_id);
                    }
                    // Замена и открывающего, и закрывающего тега в шаблоне
                    $template_processor->setValue($block->tag_with_id, '${' . $new_tag_with_id . '}', 1);
                    $template_processor->setValue(self::RESERVED_KEYWORD_CLOSING_SLASH . $block->tag_with_id, '${' . self::RESERVED_KEYWORD_CLOSING_SLASH . $new_tag_with_id . '}', 1);
                    // Обновление свойств объекта
                    $block->tag_with_id = $new_tag_with_id;
                }
                // Рекурсивно вызываем применение псевдонимов для всех дочерних блоков
                $block->ApplyAliases($template_processor);
            }
            // Во всех остальных случаях - просто меняем тег
            else {
                $new_tag_with_id = $block->tag_with_id;
                // Применение псевдонимов
                foreach ($this->aliases as $alias_key => $alias_value) {
                    $new_tag_with_id = str_replace($alias_key, $alias_value, $new_tag_with_id);
                }
                // Замена тега в шаблоне
                $template_processor->setValue($block->tag_with_id, '${' . $new_tag_with_id . '}', 1);
                // Обновление свойств объекта
                $block->tag_with_id = $new_tag_with_id;
            }
        }
    }

    /**
     * Обрабатывает блоки в шаблоне (раскрывает и заполняет циклы).
     * @param PhpWordTemplateProcessor $template_processor Класс для работы с объектами Word
     * @param array|null               $data               Массив данных, из которого брать значения переменных
     */
    protected function ProccessBlocks(PhpWordTemplateProcessor $template_processor, ?array $data): void
    {
        foreach ($this->blocks as $block) {
            if (is_a($block, self::class)) {
                /** @var self $block */
                if (is_a($block, BlockForeach::class)) {
                    /** @var BlockForeach $block */
                    /** @var array|string|null Массив элементов для заполнения foreach */
                    // Метод $this->template_processor->deleteBlock() почему-то ломает файл, поэтому использую преобразование null в пустой массив
                    $data_to_clone = self::GetVariableValue($block->foreach_array_name, $data) ?? [];
                    if (!is_array($data_to_clone)) {
                        throw new Exception("Попытка вызвать foreach не на массив! $block->foreach_array_name не является массивом");
                    }
                    // Клонирует тело цикла столько раз, сколько было данных
                    // И если количество итераций не равно нулю, то идёт заполнение данными
                    $block->CloneBlock($template_processor, count($data_to_clone));
                } elseif (is_a($block, BlockIf::class)) {
                    /** @var BlockIf $block */
                    // Если условие, указанное в теге - верно, то клонируем блок один раз. Если оно неверно - клонируем 0 раз (удаляет блок)
                    $block->CloneBlock($template_processor, $block->CheckCondition($data) ? 1 : 0);
                }
                // После того, как цикл склонировался и заполнился данными - обрабатываем все циклы внутри него
                $block->ProccessBlocks($template_processor, $data);
            }
        }
    }

    /**
     * Клонирует блок указанное число раз.
     * @param PhpWordTemplateProcessor $template_processor Класс для работы с объектами Word
     * @param int                      $count              Количество клонирований (если 0, то блок удалится)
     */
    protected function CloneBlock(PhpWordTemplateProcessor $template_processor, int $count): void
    {
        // Будем генерировать дополнительный номер для каждого блока, так как это помогает при дебаге
        /** @var string Новое названия для блока (потому что обрабатываемое название блока не проходит клонирование) */
        $temp_tag_name = 'block' . (rand() % 10000000);
        $template_processor->setValue($this->tag_with_id, '${' . $temp_tag_name . '}', 1);
        $template_processor->setValue(self::RESERVED_KEYWORD_CLOSING_SLASH . $this->tag_with_id, '${' . self::RESERVED_KEYWORD_CLOSING_SLASH . $temp_tag_name . '}', 1);

        // Если дублирование было
        if ($count > 0) {
            // Само клонирование блока
            if (!$template_processor->CustomCloneBlock($temp_tag_name, $count)) {
                throw new Exception("Клонирование блока '$this->tag_with_id' (временное название '$temp_tag_name') не удалось! Скорее всего, недоработка кода");
            }

            if (is_a($this, BlockForeach::class)) {
                // Дублируем подблоки, чтобы итерации произошли столько раз, сколько нужно
                // (Так как дублированные блоки имеют одинаковый ID)
                /** @var BlockForeach $this */
                $blocks_to_copy_template = self::GetCopyOfArray($this->blocks);
                $this->blocks = [];
                for ($i = 0; $i < $count; ++$i) {
                    $blocks_to_copy = self::GetCopyOfArray($blocks_to_copy_template);
                    // Замена псевдонима массива на конкретные ID с названием массива
                    $this->blocks = array_merge($this->blocks, $blocks_to_copy);
                    $this->ReplaceForeachIndex($template_processor, $this->foreach_array_name, $this->foreach_variable_name, $i);
                }
            }
        } else {
            // Удаляем весь блок
            if (!$template_processor->CustomRemoveBlock($temp_tag_name)) {
                throw new Exception("Удаление блока '$this->tag_with_id' (временное название '$temp_tag_name') не удалось! Скорее всего, недоработка кода");
            }
            // Очищаем массив дочерних блоков этого элемента, чтобы они не обрабатывались в рекурсивном вызове ProccessBlocks()
            $this->blocks = [];
        }
    }

    /**
     * Обрабатывает значения глобальных переменных в шаблоне.
     * @param PhpWordTemplateProcessor $template_processor Класс для работы с объектами Word
     * @param array|null               $data               Массив данных, из которого брать значения переменных
     */
    protected function ProcessGlobalValues(PhpWordTemplateProcessor $template_processor, ?array $data): void
    {
        foreach ($this->blocks as $block) {
            // Если блок - двойной, то рекурсивно заполняем сначала его
            if (is_a($block, BlockDouble::class)) {
                /** @var BlockDouble $block */
                $block->ProcessGlobalValues($template_processor, $data);
            }
            // Если переменная
            elseif (is_a($block, BlockVariable::class)) {
                /** @var BlockVariable $block */
                $tag_without_id = self::GetTagWithoutId($block->tag_with_id);
                /** @var array|string|null Значение переменной из массива $data */
                $value = $this->GetVariableValue(
                    $tag_without_id,
                    $data
                );
                // Если переменная найдена и она не строчка - выкидываем исключение
                if (isset($value) && (is_array($value) || is_object($value))) {
                    throw new Exception("Попытка установить значение переменной, которая на самом деле является массивом или объектом! Переменная: '$tag_without_id'");
                }
                // Если переменная-изображение
                if (is_a($block, BlockVariableImage::class)) {
                    /** @var BlockVariableImage $block */
                    $value = $value ?? $block->GetDefaultValue();
                    // Если значение не пустое - устанавливаем изображение
                    if (!empty($value)) {
                        // 1) Пропорционально масштабирует изображение под высоту
                        // 2) Далее, если новая ширина всё ещё меньше указанной ширины, то
                        // пропорционально масштабирует изображение по ширине
                        $template_processor->setImageValue($block->tag_with_id, [
                            'path' => $value,
                            'width' => $block->GetImageWidth(),
                            'height' => $block->GetImageHeight(),
                        ], 1);
                    }
                    // Если пустое - просто стираем блок
                    else {
                        $template_processor->setValue($block->tag_with_id, $value, 1);
                    }
                }
                // В остальных случаях
                else {
                    $template_processor->setValue($block->tag_with_id, $value ?? $block->GetDefaultValue(), 1);
                }
            }
        }
    }

    /**
     * {@inheritdoc}
     */
    public function GetCopy(): self
    {
        // Клонируем этот объект
        $result = parent::GetCopy();
        // Очищаем дочерние объекты и клонируем каждый из них рекурсивно
        $result->blocks = [];
        foreach ($this->blocks as $block) {
            $result->blocks[] = $block->GetCopy();
        }
        return $result;
    }
}
