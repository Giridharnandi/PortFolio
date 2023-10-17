<?php
// Check if the form was submitted
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Collect form data
    $fullname = $_POST['fullname'];
    $email = $_POST['email'];
    $message = $_POST['message'];

    // Prepare the data to be saved to the Excel sheet
    $data = array($fullname, $email, $message);

    // Create a new row in the Excel sheet
    $excelFile = 'form_submissions.xlsx';
    $excelData = [];
    if (file_exists($excelFile)) {
        $excelData = unserialize(file_get_contents($excelFile));
    }
    $excelData[] = $data;
    file_put_contents($excelFile, serialize($excelData));

    // Redirect the user back to the form with a success message
    header("Location: index.html?status=success");
    exit;
}
?>
