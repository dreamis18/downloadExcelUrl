# downloadExcelUrl
自动下载excel表格里的所有url资源

1、安装python

2、安装依赖
pip install openpyxl requests pathvalidate

3、执行命令
python excel_downloader.py your_file.xlsx

功能：可以自动下载表格里的链接到对应文件夹。首先按照表格的第一列内容创建若干文件夹，文件夹名和第一列每一行的内容相同，有几行就创建几个文件夹，每一行的资源下载到该行所对应的文件夹里，下载的文件重命名为该文件所在列的列头内容一样
