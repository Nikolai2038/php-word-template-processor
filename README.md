# php-word-template-processor

**EN** | [RU](README_RU.md)

## 1. Description

PHP template processor for Microsoft Word documents.

The solution was inspired by the Python package [python-docx-template](https://github.com/elapouya/python-docx-template), which in turn had a dependency on the `jinja2` template engine, which inspired Twig itself. The structure of a Word document is XML, and this allows the Twig template engine to be applied to this markup, having prepared it in advance. This led to the creation of this solution.

## 2. Installation

1. In the project, you need to install the Composer packages `phpoffice/phpword` and `symfony/twig-bundle`;
2. Then copy the file `./PhpWordTwigTemplateProcessor.php` to your project's source files in a convenient place and change the `namespace` in it for yourself;
3. Now in the required class, it is enough to import the class `PhpWordTwigTemplateProcessor` and use it (an example is given below).

## 3. Usage

Example:

```php
$file_name_template = 'templates/some.docx';
$file_name_result = 'public/some.docx';

// Open template
$templateProcessor = new PhpWordTwigTemplateProcessor($file_name_template);

// Fill template with data
$templateProcessor->fillWithData($data);

// Save result in file
$templateProcessor->saveAs($file_name_result);
```

## 4. Template rules

The rules of the written template engine can be described as follows:

- In a Word document, you can use Twig syntax:

    - Variables `{{ ... }}`;
    - Blocks `{% ... %}`;
    - Comments `{# ... #}`.

    Inside the blocks, you can use any syntax that Twig itself supports.

- Since it is convenient to place blocks in separate paragraphs, but you do not want to output the paragraphs with the blocks themselves, you can use a special syntax for ignoring the location of the block. To use it, you need to write the name of the Word XML markup tag immediately after opening the block. For example, to ignore a paragraph (Word XML tag `<w:p>`), you need to use `{%p ... %}` instead of `{% ... %}` (`{% p ... %}` is also acceptable).

    The following tags are currently supported in this way:

    - `p` - Paragraph;
    - `tr` - Table row.

    To add new tags, simply add them to the `$TAGS_TO_MOVE_OUT` array.

- Within the Twig syntax, you can use the `‘'` and `“”` quotes, which will be interpreted as `''` and `""` respectively. This is done to make it easier to enter the template in Word, since the first type of quotes is inserted there by default.

- In addition to the Twig functionality, the ability to insert images by their URI has been added. To do this, use the block:

    ```twig
    {% picture <uri> %}
    ...
    <image to be replaced>
    ...
    {% endpicture %}
    ```

    If it was not possible to get the image by URI, the block will not be output.

- There is also the ability to vertically merge cells located in one `for` loop.

    This is used in double loops located in tables, when the identifier of the outer loop is displayed in the first column.

    To use the merge, simply add the `{% vm %}` tag (from the words "vertical merge") to the desired cell.

## 5. Contribution

Feel free to contribute via [pull requests](https://github.com/Nikolai2038/php-word-template-processor/pulls) or [issues](https://github.com/Nikolai2038/php-word-template-processor/issues)!
