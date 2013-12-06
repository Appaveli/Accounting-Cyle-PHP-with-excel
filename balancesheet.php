<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>Untitled Document</title>
        <link href="balance.css" rel="stylesheet" type="text/css" />
    </head>

    <body>
        <?php
        require_once 'C:\xampp\htdocs\phpexcel\Classes\PHPExcel\IOFactory.php';

        $filename = 'accounting.xlsx';

        if (isset($_POST['btnSubmit'])) {
            $cash = $_POST['txtCash'];
            $prepaid_insurance = $_POST['txtPrepaidInsurance'];
            $supplies = $_POST['txtSupplies'];
            $equipment = $_POST['txtEquipment'];
            $accum_deprec = $_POST['txtLessDepre'];
            $accounts_payable = $_POST['txtAccountsPayable'];
            $income_tax_payable = $_POST['txtIncomeTaxPayable'];

            $common_stock = $_POST['txtCommonStock'];
            $retained_earning = $_POST['txtRetainedEarnings'];

            $objReader = PHPExcel_IOFactory::createReader('Excel2007');
            $objReader->setReadDataOnly(true);
            $objPHPExcel = $objReader->load($filename);
            $objWorksheet = $objPHPExcel->getActiveSheet();

            $new_balance_sheet = new PHPExcel_Worksheet($objPHPExcel, "Balance Sheet");
            $objPHPExcel->addSheet($new_balance_sheet, 0);
            $objPHPExcel->setActiveSheetIndex(0);
            $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Assets:');
            $objPHPExcel->getActiveSheet()->setCellValue('A2', 'Cash');
            $objPHPExcel->getActiveSheet()->setCellValue('A3', 'Prepaid Insurance');
            $objPHPExcel->getActiveSheet()->setCellValue('A4', 'Supplies');
            $objPHPExcel->getActiveSheet()->setCellValue('A5', 'Equipment');
            $objPHPExcel->getActiveSheet()->setCellValue('A6', 'Less: Accum Depre');
            $objPHPExcel->getActiveSheet()->setCellValue('A7', 'Total Assets');
            $objPHPExcel->getActiveSheet()->setCellValue('A9', 'Liabilities');
            $objPHPExcel->getActiveSheet()->setCellValue('A10', 'Accounts Payable');
            $objPHPExcel->getActiveSheet()->setCellValue('A11', 'Income Tax Payable');
            $objPHPExcel->getActiveSheet()->setCellValue('A12', 'Stockholders Equity');
            $objPHPExcel->getActiveSheet()->setCellValue('A15', 'Common Stock');
            $objPHPExcel->getActiveSheet()->setCellValue('A14', 'Retained Earnings');
            $objPHPExcel->getActiveSheet()->setCellValue('A15', 'Total Stockholders Equity');
            $objPHPExcel->getActiveSheet()->setCellValue('A16', 'Total Liabilities &');
            $objPHPExcel->getActiveSheet()->setCellValue('A17', 'Stockholders Equity');
            $objPHPExcel->getActiveSheet()->setCellValue('B2', $cash);
            $objPHPExcel->getActiveSheet()->setCellValue('B3', $prepaid_insurance);
            $objPHPExcel->getActiveSheet()->setCellValue('B4', $supplies);
            $objPHPExcel->getActiveSheet()->setCellValue('B5', $equipment);
            $objPHPExcel->getActiveSheet()->setCellValue('C6', $accum_deprec);
            $objPHPExcel->getActiveSheet()->setCellValue('C10', $accounts_payable);
            $objPHPExcel->getActiveSheet()->setCellValue('C11', $income_tax_payable);
            $objPHPExcel->getActiveSheet()->setCellValue('C13', $common_stock);
            $objPHPExcel->getActiveSheet()->setCellValue('C14', $retained_earning);
            $objPHPExcel->getActiveSheet()->setCellValue('C15', '=SUM(C13:C14)');
            $objPHPExcel->getActiveSheet()->setCellValue('C12', '=SUM(C10:C11)');
            $objPHPExcel->getActiveSheet()->setCellValue('C16', '=C12+C15');
            $objPHPExcel->getActiveSheet()->setCellValue('C7', '=(B2)+(B3)+(B4)+(B5)-(C6)');
            // $objPHPExcel ->getActiveSheet()->setCellValue('C5','=SUM(A5:B5)');
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            $objWriter->save($filename);
            echo 'sheet created';
        }
        ?>
        <form id="form1" name="form1" method="post" action="">

            <div id="header"></div>
            <div id="container">
                <div id="content">
                    <div align="center">
                        <div align="left">Assets: <br />
                            Cash<input name="txtCash" type="text" class="TextMargin" size="20" /><br /><br />
                            Prepaid Insurance<input name="txtPrepaidInsurance" type="text" class="TextMargin" size="20" /><br />
                            Supplies<input name="txtSupplies" type="text" class="TextMargin" size="20" /><br />
                            Equipment<input name="txtEquipment" type="text" class="TextMargin" size="20" /><br />
                            Less: Accum Depre <input name="txtLessDepre" type="text" class="TextMargin" /><br /><br />
                            Liabilities:<br />
                            Accounts Payable<input name="txtAccountsPayable" type="text" class="TextMargin" size="20" maxlength="20" /><br />
                            Income Tax Payable
                            <input name="txtIncomeTaxPayable" type="text" class="TextMargin" size="20" /><br />

                            Stockholders' equity <br />
                            Common Stock<input name="txtCommonStock" type="text" class="TextMargin" /><br />
                            Retained Earning: <input name="txtRetainedEarnings" type="text" class="TextMargin" /><br />
                            <input type="submit" name="btnSubmit" id="btnSubmit" value="submit" />

                        </div>
                    </div>

                </div>
            </div>


            </div>

        </form>

    </body>
</html>