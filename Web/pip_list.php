<?php
// Python 路径
$pythonPath = "C:\\Users\\Kanata\\AppData\\Local\\Programs\\Python\\Python311\\python.exe";

// 要执行的命令：列出已安装的包
$command = "$pythonPath -m pip list 2>&1";

// 执行命令并捕获输出
$output = shell_exec($command);

// 显示命令的输出
echo "<pre>$output</pre>";
?>
