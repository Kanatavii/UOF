<?php
if (!isset($_GET['file'])) {
    die('No file specified.');
}

$file = 'C:/Server/JBC/' . basename($_GET['file']);

// 检查文件是否存在
if (!file_exists($file)) {
    die('File not found.');
}

// 设置下载头信息
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . basename($file) . '"');
header('Content-Length: ' . filesize($file));

// 输出文件内容
readfile($file);
exit;
?>
