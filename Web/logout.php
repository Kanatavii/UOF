<?php
session_start();

// 结束会话
session_unset();
session_destroy();

// 重定向到登录页面
header("Location: index.php");
exit;
?>
