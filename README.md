# AutoServerProfile
自动导出成Excel
1.判断是否存在文件夹，若不存在则生成
File file = new File("D:\\文件名");
if (!file.exists()) {
	file.mkdir();//创建文件夹
}

2.文件输出流的书写
3.设置变量（本文中为时间）
4.拼接需要执行的sql语句
5.调用BaseDaoDemo中findListBySql方法，获取结果list<map>集合
6.调用poi方法，创建表格对象，传入所需的几个参数，输出结果
