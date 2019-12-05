import os
import sys
import shutil


print('程序开始运行……')


#在c盘下新建python文件夹
if os.path.exists('C:\\python') == False:
    os.makedirs('C:\\python')
    print('创建C:\\python文件夹')
    

shutil.copy('geckodriver64位v0.23.exe','C:\\python\\geckodriver.exe')
print('将浏览器驱动复制到C:\\python文件夹')


input('系统环境设置完成，按回车键退出')
    
