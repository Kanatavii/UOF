<?php
session_start();

// 如果用户已经登录，直接重定向到欢迎页面
if (isset($_SESSION['username'])) {
    header("Location: welcome.php");
    exit;
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Advanced Login Page</title>
    <style>
        /* 全局样式 */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            height: 100vh;
            background: url('background.png') no-repeat center center fixed;
            background-size: cover; /* 背景图像覆盖整个页面 */
            display: flex;
            justify-content: center;
            align-items: center;
            font-family: 'Poppins', sans-serif;
            overflow: hidden; /* 防止动画溢出 */
            position: relative;
        }

        .login-container {
            background: rgba(255, 255, 255, 0.4); /* 更加透明的背景 */
            backdrop-filter: blur(15px); /* 增强虚化效果 */
            width: 400px;
            padding: 40px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
            position: relative;
            z-index: 1; /* 确保登录框在背景之上 */
        }

        h2 {
            text-align: center;
            margin-bottom: 20px;
            color: #333;
            font-size: 24px;
        }

        .input-group {
            position: relative;
            margin-bottom: 25px;
        }

        .input-group input {
            width: 100%;
            padding: 15px;
            font-size: 16px;
            border: 2px solid #e9e9e9;
            border-radius: 10px;
            transition: 0.3s;
        }

        .input-group input:focus {
            border-color: #6e8efb;
            box-shadow: 0 0 8px rgba(110, 142, 251, 0.5);
        }

        .submit-btn {
            width: 100%;
            padding: 15px;
            background: linear-gradient(135deg, #6e8efb, #a777e3);
            color: #fff;
            border: none;
            border-radius: 10px;
            font-size: 16px;
            cursor: pointer;
            transition: 0.3s;
        }

        .submit-btn:hover {
            background: linear-gradient(135deg, #5a75e3, #9466d8);
            box-shadow: 0 4px 15px rgba(110, 142, 251, 0.5);
        }

        /* 响应式设计 */
        @media (max-width: 600px) {
            .login-container {
                width: 90%;
                padding: 30px;
            }
        }
    </style>
</head>
<body>
    <div class="login-container">
        <h2>UOF</h2> <!-- 修改标题 -->
        <form action="login.php" method="POST">
            <div class="input-group">
                <input type="text" id="username" name="username" placeholder="Username" required>
            </div>
            <div class="input-group">
                <input type="password" id="password" name="password" placeholder="Password" required>
            </div>
            <button type="submit" class="submit-btn">Login</button>
        </form>
    </div>
</body>
</html>