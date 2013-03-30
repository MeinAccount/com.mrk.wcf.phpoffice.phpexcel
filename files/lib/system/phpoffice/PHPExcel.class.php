<?php
namespace wcf\system\phpoffice;
use wcf\system\exception\SystemException;
use wcf\system\SingletonFactory;
use \PHPExcel_Exception;
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
	 * Wrapper for methods of PHPExcels IOFactory
	 * 
	 * @param	string		$function
	 * @param	array		$arguments
	 */
	public function __call($function, $arguments) {
		if (method_exists('PHPExcel_IOFactory', $function)) {
			try {
				return call_user_func_array(array('PHPExcel_IOFactory', $function), $arguments);
			} catch (PHPExcel_Exception $exception) {
				throw new SystemException($exception->getMessage(), $exception->getCode(), '', $exception);
			}
		} else {
			throw new SystemException('Can not call file method ' . $function);
		}
	}
}