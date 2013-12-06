<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>Untitled Document</title>
        <link href="retained.css" rel="stylesheet" type="text/css" />
    </head>

    <body>
        <?php
        require_once 'C:\xampp\htdocs\phpexcel\Classes\PHPExcel\IOFactory.php';

        $filename = 'accounting.xlsx';
        
        if (isset($_POST['btnSave'])) {

            $net_income = $_POST['txtNetIncome'];
            $dividend = $_POST['txtDividends'];

            $objReader = PHPExcel_IOFactory::createReader('Excel2007');
            $objReader->setReadDataOnly(true);
            $objPHPExcel = $objReader->load($filename);
            $objWorksheet = $objPHPExcel->getActiveSheet();

            $new_retained_sheet = new PHPExcel_Worksheet($objPHPExcel, "Income Statement");
            $objPHPExcel->addSheet($new_retained_sheet, 2);
            $objPHPExcel->setActiveSheetIndex(2);
            $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Retained Earnings, December 1:');
            $objPHPExcel->getActiveSheet()->setCellValue('A2', 'Add: Net Income');
            $objPHPExcel->getActiveSheet()->setCellValue('A4', 'Subtotal:');
            $objPHPExcel->getActiveSheet()->setCellValue('A5', 'Less: Dividends');
            $objPHPExcel->getActiveSheet()->setCellValue('A6', 'Retained Earnings, December 31');
            $objPHPExcel->getActiveSheet()->setCellValue('B2', $net_income);
            $objPHPExcel->getActiveSheet()->setCellValue('B5', $dividend);
            $objPHPExcel->getActiveSheet()->setCellValue('B6', '=B2 - B5');

            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            $objWriter->save($filename);  
        }
            echo 'sheet created';
        
        ?>
        <form id="form1" name="form1" method="post" action="">

            <div id="header"></div>
            <div id="container">
                <div id="content">
                    <div align="center">
                        <div align="left">Retained earnings: <br />
                            Add: Net Income<input name="txtNetIncome" type="text" class="TextMargin" size="20" /><br /><br />
                            Subtotal: <br />
                            Less: Dividends<input name="txtDividends" type="text" class="TextMargin" size="20" /><br /><br />
                            <input name="btnSave" type="submit" value="Submit" />

                        </div>
                    </div>


                </div>
        </form>

    </body>
</html>