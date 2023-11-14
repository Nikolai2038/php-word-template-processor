<?php

namespace App\Helpers\PhpWordHelper;

use App\Helpers\PhpWord\Blocks\BlockDocument;
use Exception;
use RecursiveArrayIterator;
use RecursiveIteratorIterator;

/**
 * Вспомогательный класс для работы с Word.
 */
abstract class PhpWordHelper
{
    /**
     * Скачивает файл Word, сгенерированный из шаблона.
     * @param string $template_file_path Название файла шаблона (абсолютный путь в контейнере или относительно папки public)
     * @param string $result_file_name Название файла, который будет скачан (указывать с расширением (.docx))
     * @param array|null $data Массив данных, по которым будет заполняться шаблон
     */
    public static function ExportFromTemplate(
        string $template_file_path,
        string $result_file_name,
        ?array $data
    ): void
    {
        // Проверка данных на корректность
        self::CheckData($data);

        $template_processor = new CustomTemplateProcessor($template_file_path);

        // Создаём блок документа, рекурсивно заполняя его другими блоками
        $block_document = new BlockDocument($template_processor);

        // Заполняем все блоки по переданным данным
        $block_document->FillData($data);

        // Сохраняем файл
        $template_processor->saveAs($result_file_name);
    }

    /**
     * Проверяет массив входных данных на корректность.
     * @param array|null $data Массив данных, по которым будет заполняться шаблон
     */
    private static function CheckData(?array $data): void
    {
        if (isset($data)) {
            // Рекурсивно находим все ключи массива $data (включая вложения)
            $data_keys = [];
            foreach (
                new RecursiveIteratorIterator(
                    new RecursiveArrayIterator($data),
                    RecursiveIteratorIterator::SELF_FIRST
                )
                as $key => $value
            ) {
                $data_keys[] = $key;
            }
            // Проверка на наличие зарезервированных слов в ключах $data
            foreach (Block::RESERVED_KEYWORDS as $reserved_tag) {
                if (in_array($reserved_tag, $data_keys, true)) {
                    throw new Exception("В массиве \$data нельзя использовать зарезервированный ключ '$reserved_tag'");
                }
            }
        }
    }
}
