<?php

namespace App\Helpers\PhpWord;

use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\Shared\ZipArchive;
use PhpOffice\PhpWord\TemplateProcessor;
use Twig\Environment;
use Twig\Error\LoaderError;
use Twig\Error\SyntaxError;
use Twig\Loader\ArrayLoader;

/**
 * @see TemplateProcessor
 */
class TwigTemplateProcessor extends TemplateProcessor
{
    /**
     * @throws CopyFileException
     * @throws CreateTemporaryFileException
     *
     * @noinspection PhpMissingParentConstructorInspection
     * @see          TemplateProcessor::__construct()
     */
    public function __construct($documentTemplate)
    {
        // Temporary document filename initialization
        $this->tempDocumentFilename = tempnam(Settings::getTempDir(), 'PhpWord');
        if (false === $this->tempDocumentFilename) {
            throw new CreateTemporaryFileException(); // @codeCoverageIgnore
        }

        // Template file cloning
        if (false === copy($documentTemplate, $this->tempDocumentFilename)) {
            throw new CopyFileException($documentTemplate, $this->tempDocumentFilename); // @codeCoverageIgnore
        }

        // Temporary document content extraction
        $this->zipClass = new ZipArchive();
        $this->zipClass->open($this->tempDocumentFilename);
        $index = 1;
        while (false !== $this->zipClass->locateName($this->getHeaderName($index))) {
            $this->tempDocumentHeaders[$index] = $this->readPartWithRels($this->getHeaderName($index));
            ++$index;
        }
        $index = 1;
        while (false !== $this->zipClass->locateName($this->getFooterName($index))) {
            $this->tempDocumentFooters[$index] = $this->readPartWithRels($this->getFooterName($index));
            ++$index;
        }

        $this->tempDocumentMainPart = $this->readPartWithRels($this->getMainPartName());
        $this->tempDocumentSettingsPart = $this->readPartWithRels($this->getSettingsPartName());
        $this->tempDocumentContentTypes = $this->zipClass->getFromName($this->getDocumentContentTypesName());
    }

    /**
     * @see TemplateProcessor::readPartWithRels()
     */
    protected function readPartWithRels($fileName): array|string|null
    {
        $relsFileName = $this->getRelationsName($fileName);
        $partRelations = $this->zipClass->getFromName($relsFileName);
        if ($partRelations !== false) {
            $this->tempDocumentRelations[$fileName] = $partRelations;
        }

        return $this->fixBrokenMacros($this->zipClass->getFromName($fileName));
    }

    /**
     * @see TemplateProcessor::fixBrokenMacros()
     */
    protected function fixBrokenMacros($documentPart): array|string|null
    {
        return preg_replace_callback(
            '/\{[{%]([^{}]+)[}%]}/',
            function ($match) {
                return strip_tags($match[0]);
            },
            $documentPart,
        );
    }

    /**
     * @throws LoaderError
     * @throws SyntaxError
     */
    public function fillWithData(array $data): void
    {
        $loader = new ArrayLoader();
        $twig = new Environment($loader);
        $xml_for_twig = $this->tempDocumentMainPart;
        $template = $twig->createTemplate($xml_for_twig);
        $this->tempDocumentMainPart = $template->render($data);
    }
}
