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
 * sharedStrings.xml, worksheets and styles.
 *
 * Functions of the library:
 *
 * * close - Resets the class for a new run.
 * * load - Loads a .xlsx file into memory and unzips it.
 * * replace - Does a global search and replace with the given values.
 * * save - Saves the temporary buffer to disk.
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
  private $_entries      = array();
  private $_filename     = NULL;
  private $_filepath     = NULL;
  private $_tempFilename = NULL;
  private $_tempFilepath = NULL;
  private $_zipArchive   = NULL;

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
    $this->_entries = array();
    $this->_filename = NULL;
    $this->_filepath = NULL;
    $this->_tempFilename = NULL;
    $this->_tempFilepath = NULL;
    $this->_zipArchive = NULL;
  }

  /**
   * Compiles a fresh XML document from entries.
   *
   * @param array $lines
   *   The lines to compile from.
   *
   * @return string
   *   Newly formed XML.
   */
  private function compileXML($lines) {
    $output = '';

    if (count($lines)) {
      foreach ($lines as $line) {
        if (strpos($line, '>') !== FALSE) {
          $output .= '<';
        }

        $output .= $line;
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

    $this->_tempFilename = '.' . microtime(TRUE) . '.temp.xlsx';
    $this->_tempFilepath = XLSX_BASE_PATH . $this->_tempFilename;

    copy(
      $this->_filepath,
      $this->_tempFilepath
    );

    $this->_zipArchive = new ZipArchive();
    $this->_zipArchive->open($this->_tempFilepath);

    for ($i = 0; $i < $this->_zipArchive->numFiles; $i++) {
      $stat = $this->_zipArchive->statIndex($i);

      if (substr($stat['name'], -4) == '.xml') {
        $xml = $this->_zipArchive->getFromName($stat['name']);

        if (!empty($xml)) {
          $this->_entries[] = array(
            'name'  => $stat['name'],
            'lines' => explode('<', $xml),
          );
        }
      }
    }
  }

  /**
   * Does a global search and replace with the given values.
   *
   * @param array $values
   *   A list of search and replace values.
   * @param bool $treatAsTags
   *   Check if keys in the array is wrapped in ${}.
   */
  public function replace($values, $treatAsTags = FALSE) {
    if (is_array($values) &&
      count($values)) {
      foreach ($values as $key => $value) {
        if ($treatAsTags) {
          if (strlen($key) > 1 && substr($key, 0, 2) !== '${') {
            $key = '${' . $key;
          }
          if (strlen($key) > 1 && substr($key, -1) !== '}') {
            $key .= '}';
          }
        }

        if (count($this->_entries)) {
          for ($i = 0; $i < count($this->_entries); $i++) {
            if (count($this->_entries[$i]['lines'])) {
              for ($j = 0; $j < count($this->_entries[$i]['lines']); $j++) {
                $this->_entries[$i]['lines'][$j] = str_replace($key, $value, $this->_entries[$i]['lines'][$j]);
              }
            }
          }
        }
      }
    }
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
    if (count($this->_entries)) {
      for ($i = 0; $i < count($this->_entries); $i++) {
        $this->_zipArchive->addFromString($this->_entries[$i]['name'], $this->compileXML($this->_entries[$i]['lines']));
      }
    }

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
}