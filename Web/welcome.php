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

$columns = [];
$limit = isset($_POST['limit']) ? (int)$_POST['limit'] : 25;

// 定义要隐藏的列
$hidden_columns = ['到港时间', '搬入时间', '福山', '佐川', 'TONAMI', '中村包车', 'JHSS', 'UOF混载', 'UOF包车', 'UOF包车', 'UOF包车', '成本含税', '请求含税', 'FBA', '请求含税', '请求含税']; // 替换为你要隐藏的列名

// 查询数据库中的表格信息
$sql = "SELECT * FROM zongdan_2024 LIMIT $limit";
$result = $conn->query($sql);
if ($result && $result->num_rows > 0) {
    $fields = $result->fetch_fields();
    foreach ($fields as $field) {
        if (!in_array($field->name, $hidden_columns)) {
            $columns[] = $field->name;
        }
    }
}

// 搜索功能
$search_query = "";
if (isset($_POST['search'])) {
    $search_query = $conn->real_escape_string($_POST['search']);
    if (!empty($columns)) {
        $sql = "SELECT * FROM zongdan_2024 WHERE CONCAT_WS(' ', " . implode(', ', array_map(function($col) { return "`" . $col . "`"; }, $columns)) . ") LIKE '%$search_query%' LIMIT $limit";
        $result = $conn->query($sql);
        if (!$result) {
            die("SQL 错误: " . $conn->error);
        }
    }
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Database Table Display</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="jquery-easyui/themes/default/easyui.css">
    <link rel="stylesheet" type="text/css" href="jquery-easyui/themes/icon.css">
    <script type="text/javascript" src="jquery-easyui/jquery.min.js"></script>
    <script type="text/javascript" src="jquery-easyui/jquery.easyui.min.js"></script>
    <script type="text/javascript" src="jquery-easyui/locale/easyui-lang-zh_CN.js"></script>
    <script>
        function loadTable(tableName) {
            // 修改表格的加载方式，通过改变 URL 参数来切换数据库
            window.location.href = "?table=" + tableName;
        }

        $(document).ready(function() {
            $("#filter-button").click(function() {
                var filterLicence = $("#filter-licence").prop('checked');
                var filterWarehouse = $("#filter-warehouse").prop('checked');

                $.ajax({
                    url: 'filter.php?table=<?php echo isset($_GET['table']) ? $_GET['table'] : 'osaka'; ?>',
                    type: 'POST',
                    data: {
                        filter_licence: filterLicence,
                        filter_warehouse: filterWarehouse
                    },
                    success: function(response) {
                        var data = JSON.parse(response);
                        var rows = "";

                        data.forEach(function(row) {
                            rows += "<tr>";
                            for (var key in row) {
                                rows += "<td>" + row[key] + "</td>";
                            }
                            rows += "</tr>";
                        });

                        $("tbody").html(rows);
                    },
                    error: function() {
                        alert("获取数据时出错");
                    }
                });
            });
        });
    </script>
    <style>
        body {
            font-family: 'Roboto', sans-serif;
            background-color: #e9ecef;
            margin: 0;
            padding: 0;
        }
        .easyui-layout {
            height: 100vh;
            width: 100vw;
        }
        .header {
            padding: 20px;
            background-color: #343a40;
            color: #ffffff;
            text-align: center;
            font-weight: bold;
            border-bottom: 4px solid #ffc107;
        }
        .logout-btn {
            position: absolute;
            top: 20px;
            right: 20px;
            background-color: #ffc107;
            color: #343a40;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 30px;
            font-weight: bold;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
            transition: background-color 0.3s, box-shadow 0.3s;
        }
        .logout-btn:hover {
            background-color: #e0a800;
            box-shadow: 0 6px 12px rgba(0, 0, 0, 0.2);
        }
        .content {
            padding: 20px;
        }
        .table-buttons {
            margin-bottom: 20px;
        }
        .table-btn {
            padding: 10px 20px;
            margin-right: 10px;
            background-color: #007bff;
            color: #ffffff;
            border: none;
            cursor: pointer;
            font-size: 16px;
        }
        .table-btn:hover {
            background-color: #0056b3;
        }
        .search-bar {
            margin-bottom: 20px;
        }
        .search-bar input[type="text"] {
            width: 300px;
            padding: 10px;
            font-size: 16px;
        }
        .search-bar select {
            padding: 10px;
            font-size: 16px;
            margin-left: 10px;
        }
        .search-bar button {
            padding: 10px 20px;
            font-size: 16px;
            background-color: #007bff;
            color: #ffffff;
            border: none;
            cursor: pointer;
            margin-left: 10px;
        }
        .search-bar button:hover {
            background-color: #0056b3;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        table,
        th,
        td {
            border: 1px solid #dee2e6;
        }
        th,
        td {
            padding: 12px;
            text-align: left;
            word-wrap: break-word;
            white-space: nowrap;
        }
        th {
            background-color: #007bff;
            color: #ffffff;
            text-transform: uppercase;
            font-weight: 700;
            border-bottom: 3px solid #ffc107;
        }
        tr:nth-child(even) {
            background-color: #f8f9fa;
        }
        tr:hover {
            background-color: #f1f1f1;
        }
    </style>
</head>
<body class="easyui-layout">
    <div data-options="region:'north'" class="header">
        <h1>Welcome, <?php echo htmlspecialchars($_SESSION['username']); ?>!</h1>
        <a href="logout.php" class="logout-btn"><i class="fas fa-sign-out-alt"></i> Logout</a>
    </div>

    <div data-options="region:'center'" class="content">
        <div class="table-buttons">
            <button onclick="loadTable('osaka')" class="table-btn">大阪</button>
            <button onclick="loadTable('tokyo')" class="table-btn">东京</button>
        </div>
        <div class="search-bar">
            <form method="POST" action="">
                <input type="text" name="search" value="<?php echo htmlspecialchars($search_query); ?>" placeholder="Search...">
                <label for="limit">显示行数:</label>
                <select name="limit" id="limit">
                    <option value="25" <?php echo ($limit == 25) ? 'selected' : ''; ?>>25</option>
                    <option value="50" <?php echo ($limit == 50) ? 'selected' : ''; ?>>50</option>
                    <option value="100" <?php echo ($limit == 100) ? 'selected' : ''; ?>>100</option>
                    <option value="200" <?php echo ($limit == 200) ? 'selected' : ''; ?>>200</option>
                    <option value="500" <?php echo ($limit == 500) ? 'selected' : ''; ?>>500</option>
                    <option value="1000" <?php echo ($limit == 1000) ? 'selected' : ''; ?>>1000</option>
                </select>
                <button type="submit">Search</button>
            </form>
        </div>
        <div class="filter-bar">
            <label>
                <input type="checkbox" id="filter-licence"> 许可时间不为空
            </label>
            <label>
                <input type="checkbox" id="filter-warehouse"> 入库时间为空
            </label>
            <button id="filter-button">筛选</button>
        </div>
        <table class="easyui-datagrid" data-options="rownumbers:true, singleSelect:true, fit:true, fitColumns:true">
            <thead>
                <tr>
                    <?php foreach ($columns as $column) { ?>
                        <th data-options="field:'<?php echo htmlspecialchars($column); ?>', width:200"> <?php echo htmlspecialchars($column); ?></th>
                    <?php } ?>
                </tr>
            </thead>
            <tbody>
                <?php
                // 重新执行查询以生成表格内容
                if ($result && $result->num_rows > 0) {
                    while ($row = $result->fetch_assoc()) {
                        echo '<tr>';
                        foreach ($row as $key => $cell) {
                            if (!in_array($key, $hidden_columns)) {
                                echo '<td>' . htmlspecialchars((string)$cell) . '</td>';
                            }
                        }
                        echo '</tr>';
                    }
                } else {
                    echo '<tr><td colspan="' . count($columns) . '">No data available</td></tr>';
                }
                $conn->close();
                ?>
            </tbody>
        </table>
    </div>
</body>
</html>
