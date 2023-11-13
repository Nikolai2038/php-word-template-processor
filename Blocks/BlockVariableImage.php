<?php

namespace App\Helpers\PhpWordHelper\Blocks;

use App\Helpers\PhpWordHelper\PhpWordTemplateProcessor;
use Exception;

/**
 * Блок изображения (одинарный тег).
 */
class BlockVariableImage extends BlockVariable
{
    /** @var int Количество специальных единиц масштабирования изображения.
     * Достаточно большое (пользователь вряд ли укажет такое) */
    private const UNLIMITED_SIZE_IN_IMAGE_POINTS = 2500; // Примерно 50 см

    /**
     * Ширина изображения.
     * @var int|null Ширина изображения или null, если она не установлена
     */
    private ?int $width = null;

    /**
     * Высота изображения.
     * @var int|null Высота изображения или null, если она не установлена
     */
    private ?int $height = null;

    /**
     * Создаёт блок изображения.
     * {@inheritdoc}
     */
    public function __construct(string $tag_with_id, PhpWordTemplateProcessor $template_processor)
    {
        // -------------------------------------------------------------------------------------------------
        // Сначала нужно вызвать родительский метод для установки значения по умолчанию
        // -------------------------------------------------------------------------------------------------
        // Убираем из тега префикс ключевого слова ("${image <url>[:<default>]}" -> "${<url>[:<default>]}")
        $this->tag_with_id = str_replace(self::RESERVED_KEYWORD_IMAGE . ' ', '', $tag_with_id);
        // Заменяем сам тег в шаблоне
        $template_processor->setValue($tag_with_id, '${' . $this->tag_with_id . '}', 1);
        parent::__construct($this->tag_with_id, $template_processor);
        // -------------------------------------------------------------------------------------------------

        $old_tag_with_id = $this->tag_with_id;

        // Всё до ":"
        $tag_before_default_value = explode(':', $tag_with_id, 2)[0];
        // "#<id>#image <width> <height> <url>"
        $tag_data = explode(' ', $tag_before_default_value);
        $count = count($tag_data);
        if ($count > 4) {
            throw new Exception("Неправильный синтаксис тега image! Тег '$tag_with_id'. Требуемый синтаксис: 'image[ <ширина в см>[ <высота в см>]] <url>[:<default_url>]'");
        } // Указано ключевое слово, ширина/высота и url
        elseif ($count === 4) {
            if ($tag_data[1] === 'null' && $tag_data[2] === 'null') {
                throw new Exception("Ширина и высота не могут быть одновременно равны null!");
            }
            // Установка ширины
            if ($tag_data[1] === 'null') {
                $this->width = self::UNLIMITED_SIZE_IN_IMAGE_POINTS;
            } else {
                $this->width = self::GetImageUnitsFromCentimeteres(floatval($tag_data[1]));
            }
            // Установка высоты
            if ($tag_data[2] === 'null') {
                $this->height = self::UNLIMITED_SIZE_IN_IMAGE_POINTS;
            } else {
                $this->height = self::GetImageUnitsFromCentimeteres(floatval($tag_data[2]));
            }
            // Убираем из названия тега информацию о ширине и высоте
            $this->tag_with_id = $tag_data[0] . ' ' . $tag_data[3];
        } // Если указано просто значение, то рассматриваем его как ширину
        elseif ($count === 3) {
            $this->width = self::GetImageUnitsFromCentimeteres(floatval($tag_data[1]));
            $this->height = self::UNLIMITED_SIZE_IN_IMAGE_POINTS;
            // Убираем из названия тега информацию о ширине
            $this->tag_with_id = $tag_data[0] . ' ' . $tag_data[2];
        }
        // Иначе ширина и высота не будут установлены

        $this->tag_with_id = str_replace(self::RESERVED_KEYWORD_IMAGE . ' ', '', $this->tag_with_id);
        $template_processor->setValue($old_tag_with_id, '${' . $this->tag_with_id . '}', 1);
    }

    /**
     * Возвращает ширину изображения.
     * @return int|null Ширина изображения или null, если она не установлена
     */
    public function GetImageWidth(): ?int
    {
        return $this->width;
    }

    /**
     * Возвращает высоту изображения.
     * @return int|null Высота изображения или null, если она не установлена
     */
    public function GetImageHeight(): ?int
    {
        return $this->height;
    }

    /**
     * Получает количество специальных единиц масштабирования изображения из сантиметров.
     * @param float $centimeteres Сантиметры
     * @return int   Количество специальных единиц масштабирования изображения
     */
    private static function GetImageUnitsFromCentimeteres(float $centimeteres): int
    {
        // Эту величину узнал практическим путём - без понятия, что это
        // К тому же, она немного не точна бывает - с погрешностью до +- 0.1
        return intval($centimeteres / 0.0265);
    }
}
