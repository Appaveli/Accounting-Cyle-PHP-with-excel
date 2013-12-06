<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>Untitled Document</title>
        <link href="income.css" rel="stylesheet" type="text/css" />
    </head>

    <body>
        <?php
        require_once 'C:\xampp\htdocs\phpexcel\Classes\PHPExcel\IOFactory.php';

        $filename = 'accounting.xlsx';

        if (isset($_POST['btnSave'])) {
            $landscaping_revenue = $_POST['txtLandscape'];
            $rent_expenses = $_POST['txtRentExpenses'];
            $salaries_expenses = $_POST['txtSalariesExpenses'];
            $insurance_expenses = $_POST['txtInsuranceExpenses'];
            $supplies_expenses = $_POST['txtSuppliesExpense'];
            $depreciation_expenses = $_POST['txtDepreciationExpense'];
            $income_tax_expenses = $_POST['txtIncomeTaxExpense'];

            $objReader = PHPExcel_IOFactory::createReader('Excel2007');
            $objReader->setReadDataOnly(true);
            $objPHPExcel = $objReader->load($filename);
            $objWorksheet = $objPHPExcel->getActiveSheet();

            $new_income_sheet = new PHPExcel_Worksheet($objPHPExcel, "Income Statement");
            $objPHPExcel->addSheet($new_income_sheet, 1);
            $objPHPExcel->setActiveSheetIndex(1);
            $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Revenues:');
            $objPHPExcel->getActiveSheet()->setCellValue('A2', 'Landscaping Revenues');
            $objPHPExcel->getActiveSheet()->setCellValue('A4', 'Expenses:');
            $objPHPExcel->getActiveSheet()->setCellValue('A5', 'Rent Expenses');
            $objPHPExcel->getActiveSheet()->setCellValue('A6', 'Salaries Expenses');
            $objPHPExcel->getActiveSheet()->setCellValue('A7', 'Insurance Expenses');
            $objPHPExcel->getActiveSheet()->setCellValue('A8', 'Supplies Expenses');
            $objPHPExcel->getActiveSheet()->setCellValue('A9', 'Depreciation Expenses');
            $objPHPExcel->getActiveSheet()->setCellValue('A10', 'Income Tax Expenses');
            $objPHPExcel->getActiveSheet()->setCellValue('A11', 'Total Expenses');
            $objPHPExcel->getActiveSheet()->setCellValue('A13', 'Net Income');
            $objPHPExcel->getActiveSheet()->setCellValue('C1', $landscaping_revenue);
            $objPHPExcel->getActiveSheet()->setCellValue('B5', $rent_expenses);
            $objPHPExcel->getActiveSheet()->setCellValue('B6', $salaries_expenses);
            $objPHPExcel->getActiveSheet()->setCellValue('B7', $insurance_expenses);
            $objPHPExcel->getActiveSheet()->setCellValue('B8', $supplies_expenses);
            $objPHPExcel->getActiveSheet()->setCellValue('B9', $depreciation_expenses);
            $objPHPExcel->getActiveSheet()->setCellValue('B10', $income_tax_expenses);
            $objPHPExcel->getActiveSheet()->setCellValue('C11', '=SUM(B5:B10)');
            $objPHPExcel->getActiveSheet()->setCellValue('C13', '=C1 - C11');

            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            $objWriter->save($filename);
            echo 'sheet created';
        }
        ?>
        <form id="form1" name="form1" method="post" action="">
            <div id="header"></div>
            <div id="container">
                <div id="content">
                    <div align="left">Revenues: <br />
                        Landscaping Revenue<input name="txtLandscape" type="text" class="TextMargin" size="20" /><br /><br />
                        Expenses:<br />
                        Rent Expenses<input name="txtRentExpenses" type="text" class="TextMargin" size="20" maxlength="20" /><br />
                        Salaries Expenses 
                        <input name="txtSalariesExpenses" type="text" class="TextMargin" size="20" /><br />
                        Insurance Expenses <input name="txtInsuranceExpenses" type="text" class="TextMargin" /><br />
                        Supplies Expense <input name="txtSuppliesExpense" type="text" class="TextMargin" /><br />
                        Depreciation Expense <input name="txtDepreciationExpense" type="text" class="TextMargin" /><br />
                        Income Tax Expenses <input name="txtIncomeTaxExpense" type="text" class="TextMargin" /><br />
                        <input name="btnSave" type="submit" value="Submit" />

                    </div>
                </div>


            </div>

        </form>

    </body>
</html>