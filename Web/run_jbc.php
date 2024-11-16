<?php
// 设置 Python 和 jbc.py 的路径
$python = 'C:\\Users\\Kanata\\AppData\\Local\\Programs\\Python\\Python311\\python.exe';
$script = 'C:\\xampp\\htdocs\\jbc.py';

header('Content-Type: application/json'); // 返回 JSON 格式

// 执行 Python 脚本并捕获输出和错误信息
$command = escapeshellcmd("$python $script 2>&1");
$output = shell_exec($command);

// 检查 Python 脚本是否成功运行
if ($output === null) {
    echo json_encode([
        'status' => 'error',
        'message' => '无法运行 Python 脚本，请检查路径或权限。',
    ]);
    exit;
}

// 查找最新生成的文件
$directory = 'C:/Server/JBC/';
$files = glob($directory . '*.xlsx');
if (!$files) {
    echo json_encode([
        'status' => 'error',
        'message' => '未找到生成的文件。',
    ]);
    exit;
}

// 按创建时间排序，获取最新的文件
usort($files, function($a, $b) {
    return filemtime($b) - filemtime($a);
});
$latest_file = basename($files[0]);

echo json_encode([
    'status' => 'success',
    'message' => 'Python 脚本运行成功。',
    'file' => $latest_file,
]);
