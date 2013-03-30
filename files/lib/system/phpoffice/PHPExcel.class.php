<?php
namespace wcf\system\phpoffice;
use wcf\system\exception\SystemException;
use wcf\system\SingletonFactory;
use \PHPExcel_Exception;
use \PHPExcel_IOFactory;
use \PHPExcel_Settings;

/**
 * PHPExcel implementation
 * 
 * @author	Magnus Kühn
 * @copyright	2013 Magnus Kühn
 * @license	GNU Lesser General Public License <http://opensource.org/licenses/lgpl-license.php>
 * @package	com.mrk.wcf.phpoffice.phpexcel
 */
class PHPExcel extends SingletonFactory {
	/**
	 * @see wcf\system\SingletonFactory::init()
	 */
	protected function init() {
		require_once(WCF_DIR.'lib/system/api/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php');
		
		// init tcpdf libary for pdf file generation
		PHPExcel_Settings::setPdfRenderer(PHPExcel_Settings::PDF_RENDERER_TCPDF, WCF_DIR.'lib/system/api/tcpdf/');
	}
	
	/**
	 * Opens a file in PHPExcel
	 * 
	 * @param	string	$filename
	 * @return	PHPExcel
	 */
	public function open($filename) {
		try {
			return PHPExcel_IOFactory::load($filename);
		} catch (PHPExcel_Exception $exception) {
			$this->handleException($exception);
		}
	}
	
	/**
	 * Saves a file
	 * 
	 * @param	PHPExcel	$file
	 * @param	string		$filename
	 * @param	string		$writerType
	 */
	public function save(PHPExcel $file, $filename, $writerType = 'Excel2007') {
		try {
			$writer = PHPExcel_IOFactory::createWriter($file, $writerType);
			$writer->save($filename);
		} catch (PHPExcel_Exception $exception) {
			$this->handleException($exception);
		}
	}
	
	/**
	 * Handles an exception from PHPExcel
	 * 
	 * @param	PHPExcel_Exception	$exception
	 * @throws	wcf\system\exception\SystemException
	 */
	protected function handleException(PHPExcel_Exception $exception) {
		throw new SystemException($exception->getMessage(), $exception->getCode(), '', $exception);
	}
}