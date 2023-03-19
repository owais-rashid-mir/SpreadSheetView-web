<html>

<head>
    <!-- Linking CSS File.  -->
    <link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>
<form action="<?php echo $_SERVER['PHP_SELF']; ?>" method="post" enctype="multipart/form-data" style="text-align: center;">
        <!-- <p>Open a Spreadsheet file and click on the "Upload and Display" button to view the file.</p> -->
        <p style="font-size: 1.2em; font-weight: bold; text-align: center;">
         Select a Spreadsheet file and click on the "<span style="color: blue;">Upload and Display</span>" button to view the file.
        </p>
        <input type="file" name="fileToUpload" id="fileToUpload">
        <br>
        <input type="submit" value="Upload and Display" name="submit"> 
    </form> <br>

    <?php
        if(isset($_POST["submit"])) {
            // Using PhpExcel library.
            require_once "Classes/PHPExcel.php";

            //Getting file details and loading it in our project
            $fileName = $_FILES["fileToUpload"]["name"];
            $fileTmpName = $_FILES["fileToUpload"]["tmp_name"];
            $fileType = $_FILES["fileToUpload"]["type"];
            $fileSize = $_FILES["fileToUpload"]["size"];
            $fileError = $_FILES["fileToUpload"]["error"];

            //Checking if uploaded file is a valid spreadsheet file
            $allowedExtensions = array("xlsx", "xls", "csv");
            $fileExtension = strtolower(pathinfo($fileName, PATHINFO_EXTENSION));
            if(in_array($fileExtension, $allowedExtensions)){
                if($fileError === 0){
                    if($fileSize < 1000000){ //Limiting file size to 1 MB
                        //Loading the excel file
                        $reader = PHPExcel_IOFactory::createReaderForFile($fileTmpName);
                        $excel_obj = $reader->load($fileTmpName);

                        // Fetching the sheet of excel which is to be displayed. 18 is the LTB sheet number.
                        $worksheet = $excel_obj->getSheet();

                        //Printing specific values
                        //echo $worksheet->getCell('A1')->getValue();

                        // Printing all rows and columns of the excel sheet
                        
                        echo "<div style='text-align:center;'><strong>Number of rows and columns in this SpreadSheet: </strong>";
                        $lastRow = $worksheet->getHighestRow();
                        $columncount = $worksheet->getHighestDataColumn();
                        $columncount_number = PHPExcel_Cell::columnIndexFromString($columncount);

                        echo $lastRow . '   ';
                        echo $columncount;

                        echo "<table border='1' >";
                        for ($row = 0; $row <= $lastRow; $row++) {
                            echo "<tr>";
                            for ($col = 0; $col <= $columncount_number; $col++) {
                                echo "<td>";

                                /* getFormattedValue() automatically reads the calculated values. Like, if Sum function
                                is applied to a column in an excel sheet, PhpExcel will display the calculated
                                results on a webpage. */
                                echo $worksheet->getCell(PHPExcel_Cell::stringFromColumnIndex($col) . $row)->getFormattedValue();

                                echo "</td>";
                            }
                            echo "</tr>";
                        }
                        echo "</table>";
                    } else {
                        echo "File size is too large. Please upload a file less than 1 MB.";
                    }
                } else {
                    echo "There was an error uploading your file. Please try again.";
                }
            } else {
                echo "Invalid file type. Please upload a valid spreadsheet file (xlsx, xls, or csv).";
            }
        }
    ?>
</body>
</html>
