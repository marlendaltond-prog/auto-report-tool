@echo off
chcp 936 >nul
echo 简单测试脚本 > test_output.txt 2>&1
echo 当前目录：>> test_output.txt 2>&1
cd >> test_output.txt 2>&1
echo. >> test_output.txt 2>&1
echo 文件列表：>> test_output.txt 2>&1
dir >> test_output.txt 2>&1
echo. >> test_output.txt 2>&1
echo Python检查：>> test_output.txt 2>&1
where python >> test_output.txt 2>&1
echo. >> test_output.txt 2>&1
echo 运行Python测试：>> test_output.txt 2>&1
python -c "print('Python测试成功')" >> test_output.txt 2>&1
echo Python退出码：%ERRORLEVEL% >> test_output.txt 2>&1
echo. >> test_output.txt 2>&1
echo 测试完成 >> test_output.txt 2>&1