<?php
namespace wcf\system\phpoffice;
use wcf\system\SingletonFactory;

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
	}
}