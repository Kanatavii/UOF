<?php
session_start();

// 加载用户信息
$users = include('users.php');

// 获取表单数据
$username = $_POST['username'] ?? '';
$password = $_POST['password'] ?? '';

// 输入验证
if (empty($username) || empty($password)) {
    echo "Username and password are required.";
    exit;
}

// 检查用户名和密码
if (isset($users[$username]) && $users[$username] === $password) {
    // 密码正确，设置会话变量
    $_SESSION['username'] = $username;
    header("Location: welcome.php"); // 登录成功后重定向到欢迎页面
    exit;
} else {
    // 密码不正确
    echo "Invalid username or password.";
}
?>
