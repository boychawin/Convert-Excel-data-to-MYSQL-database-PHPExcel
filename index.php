<?php
$db = mysqli_connect("localhost", "root", "", "db_excel");
$output = '';
if (isset($_POST["import"])) {
    $tmp = explode(".", $_FILES["excel"]["name"]);
    $extension = end($tmp);

    $allowed_extension = array("xls", "xlsx", "csv"); //นามสกุลไฟล์ ที่อนุญาต
    if (in_array($extension, $allowed_extension)) //ตรวจสอบนามสกุลไฟล์
    {
        $file = $_FILES["excel"]["tmp_name"]; // ที่มาของไฟล์ excel
        include("PHPExcel/Classes/PHPExcel/IOFactory.php"); // เพิ่ม PHPExcel Library 
        $objPHPExcel = PHPExcel_IOFactory::load($file); // สร้างวัตถุของไลบรารี PHPExcel โดยใช้วิธี load () และในวิธีการโหลดกำหนดเส้นทางของไฟล์ที่เลือก

        $output .= "<table class='table table-striped table-hover '>";
        foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
            $highestRow = $worksheet->getHighestRow();
            for ($row = 2; $row <= $highestRow; $row++) { //สามารถปรับ $row เพื่อเลือกแถวที่บัทึกลงฐานข้อมูลได้เลย ถ้าไม่ปรับมันจะเอาข้อมูลทั้งหมดที่เห็นในตารางลงฐานข้อมูลนะครับ
                $output .= "<tr>";
                $order = mysqli_real_escape_string($db, $worksheet->getCellByColumnAndRow(0, $row)->getValue()); //ลำดับ
                $id = mysqli_real_escape_string($db, $worksheet->getCellByColumnAndRow(1, $row)->getValue()); //รหัสจังหวัด
                $name_th = mysqli_real_escape_string($db, $worksheet->getCellByColumnAndRow(2, $row)->getValue()); //ชื่อจังหวัด
                $name_en = mysqli_real_escape_string($db, $worksheet->getCellByColumnAndRow(3, $row)->getValue()); //ชื่อจังหวัดภาษาอังกฤษ

                $query = "INSERT INTO `tb_excel_data`(`name_th`, `name_en`) VALUES ('$name_th', '$name_en')"; //ผมจะ insert ข้อมูลแค่ 2 ฟิลด์

                $result = mysqli_query($db, $query);


                $output .= '<td>' . $order . '</td>';
                $output .= '<td>' . $id . '</td>';
                $output .= '<td>' . $name_th . '</td>';
                $output .= '<td>' . $name_en . '</td>';
                $output .= '</tr>';
            }
        }

        $output .= '</table>';
        if ($result) {
            echo '<div class="alert alert-success" role="alert">
            บันทึกข้อมูลสำเร็จ
          </div>';
        } else {
            echo '<div class="alert alert-danger" role="alert">
            บันทึกข้อมูลไม่สำเร็จ
          </div>';
        }
    } else {
        $output = '<div class="alert alert-danger" role="alert">
        ไฟล์ไม่ถูกต้อง
      </div>';
    }
}
?>



<!DOCTYPE html>
<html lang="th">

<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>แปลงข้อมูลใน excel ไปเก็บในฐานข้อมูล MYSQL | boychawin.com</title>
    <link rel="icon" href="https://boychawin.com/B_images/logoboychawin.com.png" type="image/png" />

    <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>
    <!-- boychawin.com -->
    <!-- <link href="https://boychawin.com/_next/static/css/d14dc5e59bd60eaeb5ad.css" rel='stylesheet'> -->
    <!-- bootstrap 5 -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-eOJMYsd53ii+scO/bJGFsiCZc+5NDVN2yr8+0RDqr0Ql0h+rP48ckxlpbzKgwra6" crossorigin="anonymous">

</head>

<body background="https://boychawin.com/B_images/logoboychawins.com.png">

    <div class="container ">
        <h3 align="center" class="mt-5">แปลงข้อมูลใน excel ไปเก็บในฐานข้อมูล MYSQL</h3><br />
        <form method="post" enctype="multipart/form-data">

            <div class="input-group">
                <input type="file" name="excel" class="form-control" id="inputGroupFile04" aria-describedby="inputGroupFileAddon04" aria-label="Upload">
                <button type="submit" name="import" class="btn btn-outline-secondary" type="button" id="inputGroupFileAddon04">เพิ่ม</button>
            </div>

        </form>
        <br />
        <br />
        <?php
        echo $output;
        ?>
    </div>

</body>

</html>