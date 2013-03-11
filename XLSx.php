<?php

/**
 * @file
 * XLSx is a PHP library designed to manipulate the content of .xlsx files.
 *
 * Maintainer: Stian Hanger (pdnagilum@gmail.com)
 *
 * .xlsx files (Microsoft Excel 2007 and newer) are basically zipped archives
 * of .xml files which holds different content. If you unzip a .xlsx file you
 * will get a lot of files, among them a xl folder containing
 * sharedStrings.xml, worksheets and styles. The sharedStrings.xml file is the
 * one we manipulated with this library.
 *
 * Functions of the library:
 *
 * * close - Resets the class for a new run.
 * * load - Loads a .xlsx file into memory and unzips it.
 * * save - Saves the temporary buffer to disk.
 * * setValue - Does a global search and replace with the given values.
 * * setValues - Does a global search and replace with an array of values.
 */

/**
 * Define the base path of the library.
 *
 * This will be used as a point of reference for creating temporary files
 */
define('XLSX_BASE_PATH', dirname(__FILE__) . '/');

/**
 * The DOCx class.
 */
class XLSx {
  private $_filename         = NULL;
  private $_filepath         = NULL;
  private $_tempFilename     = NULL;
  private $_tempFilepath     = NULL;
  private $_zipArchive       = NULL;
  private $_sharedStrings    = array();
  private $_sharedStringsXML = NULL;

  /**
   * Initiate a new instance of the XLSx class.
   *
   * @param string $filepath
   *   The .xlsx file to load.
   */
  public function __construct($filepath = NULL) {
    if ($filepath !== NULL) {
      $this->load($filepath);
    }
  }

  /**
   * Resets the class for a new run.
   */
  public function close() {
    $this->_filename = NULL;
    $this->_filepath = NULL;
    $this->_tempFilename = NULL;
    $this->_tempFilepath = NULL;
    $this->_zipArchive = NULL;
  }

  /**
   * Compiles a fresh XML document from entries.
   *
   * @param array $entries
   *   The entries to compile from.
   *
   * @return string
   *   Newly formed XML.
   */
  private function compileXML($xml) {
    $output = '';

    if (count($this->_sharedStrings)) {
      foreach ($this->_sharedStrings as $string) {
        if (strpos($string, '>') !== FALSE) {
          $output .= '<';
        }

        $output .= $string;
      }
    }

    return $output;
  }

  /**
   * Loads a .xlsx file into memory and unzips it.
   *
   * @param string $filename
   *   The .xlsx file to load.
   */
  public function load($filepath) {
    $this->_filename = (strpos($filepath, '/') !== FALSE ? substr($filepath, strrpos($filepath, '/') + 1) : $filepath);
    $this->_filepath = $filepath;

    $this->_tempFilename = '.' . time() . '.temp.xlsx';
    $this->_tempFilepath = XLSX_BASE_PATH . $this->_tempFilename;

    copy(
      $this->_filepath,
      $this->_tempFilepath
    );

    $this->_zipArchive = new ZipArchive();
    $this->_zipArchive->open($this->_tempFilepath);

    $this->_sharedStringsXML = $this->_zipArchive->getFromName('xl/sharedStrings.xml');
    $this->_sharedStrings = explode('<', $this->_sharedStringsXML);
  }

  /**
   * Saves the temporary buffer to disk.
   *
   * @param string $filepath
   *   The file to save to. If none is given the temp file is used.
   *
   * @return string
   *   The filepath of the save file.
   */
  public function save($filepath = NULL) {
    $this->_zipArchive->addFromString('xl/sharedStrings.xml', $this->compileXML($this->_sharedStrings));
    $this->_zipArchive->close();

    if ($filepath !== NULL) {
      copy(
        $this->_tempFilepath,
        $filepath
      );

      return $filepath;
    }
    else {
      return $this->_tempFilepath;
    }
  }

  /**
   * Does a global search and replace with the given values.
   *
   * @param string $search
   *   The tag to search for, represented as ${TAGNAME} in the file.
   * @param string $replace
   *   The text to replace it with.
   */
  public function setValue($search, $replace) {
    if (strlen($search) > 2 &&
      substr($search, 0, 2) !== '${' &&
      substr($search, -1) !== '}') {
      $search = '${' . $search . '}';
    }

    if (count($this->_sharedStrings)) {
      for ($i = 0; $i < count($this->_sharedStrings); $i++) {
        $this->_sharedStrings[$i] = str_replace($search, $replace, $this->_sharedStrings[$i]);
      }
    }
  }

  /**
   * Does a global search and replace with an array of values.
   *
   * @param array $values
   *   A keyed array with search and replaces values.
   */
  public function setValues($values) {
    if (is_array($values) &&
      count($values)) {
      foreach ($values as $key => $value) {
        $this->setValue($key, $value);
      }
    }
  }
}
