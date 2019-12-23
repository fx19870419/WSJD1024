import os
import sys
import shutil


print('程序开始运行……')


#在c盘下新建python文件夹
if os.path.exists('C:\\python') == False:
    os.makedirs('C:\\python')
    print('创建C:\\python文件夹')
    

shutil.copy('geckodriver32位v0.19.exe','C:\\python\\geckodriver.exe')
print('将火狐浏览器的驱动复制到C:\\python文件夹')


input('系统环境设置完成，按回车键退出')
    
