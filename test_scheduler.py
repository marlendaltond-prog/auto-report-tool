import sys
import os
import time

# 添加项目路径到系统路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from auto_report import ScheduleManager, ReportConfig, DataSourceConfig

# 创建测试用的配置
report_config = ReportConfig(
    report_name="测试报表",
    data_sources=[DataSourceConfig(
        type="excel",
        path="test.xlsx",
        parameters={"sheet_name": "Sheet1"}
    )],
    output_format=["excel"],
    calculations=[]
)

def test_schedule_manager():
    """测试ScheduleManager类的功能"""
    print("=== 测试ScheduleManager功能 ===")
    
    # 测试1: 初始化ScheduleManager
    print("\n1. 初始化ScheduleManager...")
    try:
        scheduler = ScheduleManager()
        print("✓ 成功初始化ScheduleManager")
    except Exception as e:
        print(f"✗ 初始化失败: {e}")
        return False
    
    # 测试2: 添加任务
    print("\n2. 添加调度任务...")
    try:
        task_id = scheduler.add_task(
            report_name="测试报表",
            report_config=report_config,
            schedule="* * * * *",  # 每分钟执行一次
            email_recipients=["test@example.com"]
        )
        print(f"✓ 成功添加任务，任务ID: {task_id}")
    except Exception as e:
        print(f"✗ 添加任务失败: {e}")
        return False
    
    # 测试3: 获取所有任务
    print("\n3. 获取所有调度任务...")
    try:
        tasks = scheduler.get_all_tasks()
        print(f"✓ 成功获取任务列表，共 {len(tasks)} 个任务")
        for task in tasks:
            print(f"   - 任务ID: {task['task_id']}, 报表名称: {task['report_name']}, 调度: {task['schedule']}")
    except Exception as e:
        print(f"✗ 获取任务列表失败: {e}")
        return False
    
    # 测试4: 启动调度器
    print("\n4. 启动调度器...")
    try:
        result = scheduler.start_scheduler()
        if result:
            print("✓ 成功启动调度器")
        else:
            print("! 调度器已经在运行中")
    except Exception as e:
        print(f"✗ 启动调度器失败: {e}")
        return False
    
    # 测试5: 检查调度器状态
    print("\n5. 检查调度器状态...")
    try:
        is_running = scheduler.is_running()
        print(f"✓ 调度器状态: {'运行中' if is_running else '已停止'}")
    except Exception as e:
        print(f"✗ 检查调度器状态失败: {e}")
        return False
    
    # 测试6: 停止调度器
    print("\n6. 停止调度器...")
    try:
        result = scheduler.stop_scheduler()
        if result:
            print("✓ 成功停止调度器")
        else:
            print("! 调度器已经停止")
    except Exception as e:
        print(f"✗ 停止调度器失败: {e}")
        return False
    
    # 测试7: 再次检查调度器状态
    print("\n7. 再次检查调度器状态...")
    try:
        is_running = scheduler.is_running()
        print(f"✓ 调度器状态: {'运行中' if is_running else '已停止'}")
    except Exception as e:
        print(f"✗ 检查调度器状态失败: {e}")
        return False
    
    print("\n=== 所有测试完成 ===")
    return True

if __name__ == "__main__":
    success = test_schedule_manager()
    if success:
        print("\n✅ 所有测试通过！")
        sys.exit(0)
    else:
        print("\n❌ 部分测试失败！")
        sys.exit(1)