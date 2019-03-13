<?php
namespace moxuandi\phpSpreadsheet;

use Yii;
use yii\base\InvalidConfigException;
use yii\base\Widget;
use yii\helpers\ArrayHelper;
use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * 基于 PhpSpreadsheet, 用于读取 Excel 表格的内容.
 *
 * 注意: 本类仅返回一个数组, 不处理具体导入到数据库的操作.
 *
 * @author  zhangmoxuan <1104984259@qq.com>
 * @link  http://www.zhangmoxuan.com
 * @QQ  1104984259
 * @Date  2018-9-15
 * @see https://github.com/PHPOffice/PhpSpreadsheet
 */
class ImportFile extends Widget
{
    /**
     * @var string 导入的文件路径
     */
    public $file;
    /**
     * @var bool 是否将 Excel 文件中的第一行记录设置为每行数据的键; 为`false`时将使用字母列(eg: A,B,C).
     */
    public $setFirstRecordAsKeys = true;
    /**
     * @var bool 如果 Excel 文件中有多个工作表, 是否以表名(eg:sheet1,sheet2)作为键名; 为 false 时使用数字(eg:0,1,2).
     */
    public $setIndexSheetByName = true;
    /**
     * @var string|array 当 Excel 文件中有多个工作表时, 指定仅获取某个工作表(eg:sheet1)或某几个工作表(eg:[sheet1,sheet2]).
     * 该属性为数组时, 元素的类型必须与`$setIndexSheetByName`一致:
     * `$setIndexSheetByName`为`true`时, 元素为字符串; `$setIndexSheetByName`为`false`时, 元素为整数.
     */
    public $getOnlySheet;


    public function run()
    {
        if(is_array($this->file)){
            throw new InvalidConfigException('暂不支持同时导入多个 Excel 文件！');
        }
        return self::import();
    }

    /**
     * 导入操作
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    private function import()
    {
        $spreadsheet = IOFactory::load($this->file);  // 载入 Excel 表格
        $sheetCount = $spreadsheet->getSheetCount();  // 获取 Excel 中工作表的数量
        $sheetData = [];
        if(is_string($this->getOnlySheet)){
            $sheetData = $spreadsheet->getSheetByName($this->getOnlySheet)->toArray(null, true, true, true);  // 返回一个二维数组, 数组键是行的id, 子数组键是列的大写英文字母
            if($this->setFirstRecordAsKeys){
                $sheetData = self::setFirstRecordAsLabel($sheetData);  // 将第一行记录设置为每行数据的键
            }
        }elseif($sheetCount <= 1){
            $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);  // 返回一个二维数组, 数组键是行的id, 子数组键是列的大写英文字母
            if($this->setFirstRecordAsKeys){
                $sheetData = self::setFirstRecordAsLabel($sheetData);  // 将第一行记录设置为每行数据的键
            }
        }else{
            foreach($spreadsheet->getSheetNames() as $sheetIndex => $sheetName){
                if($this->setIndexSheetByName){
                    $indexed = $sheetName;  // 索引类型: 表名索引
                    $sheet = $spreadsheet->getSheetByName($indexed);  // 按表名获取表
                }else{
                    $indexed = $sheetIndex;  // 索引类型: 数字索引
                    $sheet = $spreadsheet->getSheet($indexed);  // 按索引获取工作表
                }
                if(is_array($this->getOnlySheet) && !in_array($indexed, $this->getOnlySheet, true)){
                    continue;  // 如果不在数组`$getOnlySheet`中, 则跳过当前循环
                }
                $sheetData[$indexed] = $sheet->toArray(null, true, true, true);  // 返回一个二维数组, 数组键是行的id, 子数组键是列的大写英文字母
                if($this->setFirstRecordAsKeys){
                    $sheetData[$indexed] = self::setFirstRecordAsLabel($sheetData[$indexed]);  // 将第一行记录设置为每行数据的键
                }
            }
        }
        return $sheetData;

        //$spreadsheet = IOFactory::load($this->file);  // 载入 Excel 表格
        //$spreadsheet->getSheetCount();  // 获取 Excel 中工作表的数量
        //$spreadsheet->getSheetNames();  // 获取 Excel 中工作表的名称的列表
        //$spreadsheet->getSheet(0);  // 按索引获取工作表
        //$spreadsheet->getSheetByName('Worksheet');  // 按名称获取工作表
    }

    /**
     * 将第一行记录设置为每行数据的键, 然后返回新数组.
     * @param array $sheetData
     * @return array
     */
    private function setFirstRecordAsLabel($sheetData)
    {
        $keys = ArrayHelper::remove($sheetData, 1);  // 从数组移除第一行并返回该行的值
        $newData = [];
        foreach($sheetData as $data){
            $newData[] = array_combine($keys, $data);  // 合并两个数组来创建一个新数组, $keys为键名, $v为键值
        }
        return $newData;
    }

    /**
     * 解决导入时无法返回数组的错误.
     * @param array $config
     * @return array
     */
    public static function widget($config = [])
    {
        //return parent::widget($config);
        $config['class'] = get_called_class();
        $widget = Yii::createObject($config);
        return $widget->run();
    }
}
