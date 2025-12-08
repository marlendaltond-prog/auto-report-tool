print("Python is working")
print(f"Python version: {__import__('sys').version}")
print(f"Current directory: {__import__('os').getcwd()}")
print(f"Files in current directory: {__import__('os').listdir('.')}")