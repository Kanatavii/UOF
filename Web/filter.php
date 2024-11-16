<?php
session_start();
error_reporting(E_ALL);
ini_set('display_errors', 1);

// 如果用户未登录，重定向到登录页面
if (!isset($_SESSION['username'])) {
    header("Location: index.php");
    exit;
}

// 根据选择的表格按钮，动态设置数据库名称
if (isset($_GET['table']) && $_GET['table'] === 'tokyo') {
    $dbname = "tokyo_main";
} else {
    $dbname = "osaka_main";
}

// 数据库连接配置
$servername = "localhost";
$username = "root";
$password = "";

// 创建数据库连接
$conn = new mysqli($servername, $username, $password, $dbname);

// 检查连接是否成功
if ($conn->connect_error) {
    die("数据库连接失败: " . $conn->connect_error);
}

// 获取筛选参数
$filter_licence = isset($_POST['filter_licence']) ? $_POST['filter_licence'] : false;
$filter_warehouse = isset($_POST['filter_warehouse']) ? $_POST['filter_warehouse'] : false;

// 构建查询语句
$sql = "SELECT * FROM zongdan_2024 WHERE 1=1";

if ($filter_licence) {
    $sql .= " AND `许可时间` IS NOT NULL AND `许可时间` != ''";
}

if ($filter_warehouse) {
    $sql .= " AND (`入库时间` IS NULL OR `入库时间` = '')";
}

$result = $conn->query($sql);

if (!$result) {
    die("SQL 错误: " . $conn->error);
}

// 将结果转换为 JSON 格式
$data = [];
if ($result->num_rows > 0) {
    while ($row = $result->fetch_assoc()) {
        $data[] = $row;
    }
}

// 返回 JSON 数据
header('Content-Type: application/json');
echo json_encode($data);

$conn->close();
?>
