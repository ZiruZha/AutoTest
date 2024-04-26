高压线路继电保护自动测试软件说明
一、版本及更新说明
	1.1 20240426----->第一版----->myGui_V102.exe
		更新说明：无
	1.2 20XXXXXX----->第二版----->XXXXXXXXXX.exe
		更新说明：XXXXXX
二、简介
	2.1 背景
		高压线路保护装置选配型号多，工程发布工作量大
	2.2 目标
		自动测试，解放双手
	2.3 方法
		基于python，写一个具有GUI的自动测试软件
	2.4 思路
		测试需要使用onlly.exe；arptools；mmsdbg.exe等外部程序，该程序无现成接口可供调用，因此使用图像识别+快捷键方式实现
		主程序为GUI程序，通过button调用三个子程序，分别实现joi自动打包、保护功能自动测试、61850自动对点功能
		打包为.exe文件，无需安装python即可使用
三、功能说明/使用说明
	3.1 自动打包
		3.1.1 使用的外部软件：encryption-new.exe；ArpFilePack.exe————已经加入".\dataFiles"中
		3.1.2 使用的外部文件：《工程程序需求统计和交付管理.xlsx》————已经加入".\dataFiles"中
		3.1.3 运行步骤：
			读取《工程程序需求统计和交付管理.xlsx》中待打包joi信息
			根据装置型号选择对应的满配型号程序作为母版，将母版复制到桌面
			修改config.txt；configured.cid；configured.ccd中相关信息
			使用encryption.exe加密文件
			使用ArpFilePack.exe将已加密文件打包
		3.1.4 使用说明
			只适用NSR-303A-GCN系列程序
			使用前应将encryption-new.exe、ArpFilePack.exe目录添加到环境变量，添加方法参考https://blog.csdn.net/qq_42535133/article/details/105373924
			《工程程序需求统计和交付管理.xlsx》中包含joi打包所需的a.装置型号、b.时间、c.程序版本号、d.CQ号信息，若缺少不能打包
			满配型号程序母版在.\dataFiles\JoiMakerPrototype\NSR-303A-中，母版文件可以更新，但文件名、文件个数不能变，否则无法打包
			使用时输入起始行序号、终止行序号，点击“开始打包”
			程序运行过程中勿进行任何操作
	3.2 自动测试
		3.2.1 使用的外部软件：onlly.exe；arptools；word.exe
		3.2.2 使用的外部文件：.\\dataFiles\\《TestItems.xlsx》；.\\dataFiles\\*.png；昂立测试仪测例————已经加入".\dataFiles"中
		3.2.3 运行步骤
				按顺序测试装置差动、零序过流、距离、重合闸、选配功能并记录在word中，若无该功能则跳过
		3.2.4 使用说明
			3.2.4.1 昂立测试仪设置
				昂立测试仪软件onlly.exe需自行安装
				将测试用例复制到昂立测试仪软件目录下，dest_path = ".\MQ2自动测试20231116\param_files\通用测试\状态序列（支持递变，6U6I）"
				测试开始前手动打开OnllyMain.exe-->状态序列（支持递变，6U6I）
			3.2.4.2 arptools设置
				测试开始前打开arptools软件，且只能打开单个arptools软件
				手动连接IecDbg，所有条目文字不能遮挡
				仅支持自环测试，需要手动修改“通道识别码”、“II段保护闭锁重合闸”置0
			3.2.4.3 word设置
				需要设置二级标题快捷键为'alt'+'2'；三级标题快捷键为'alt'+'3'；四级标题快捷键为'alt'+'4'
				测试开始前打开待记录word文档，且只能打开单个word
			3.2.4.4 其他设置
				勾选装置型号
				*.png文件需根据电脑显示分辨率重新截图
				《TestItems.xlsx》禁止修改，若要变更测试项目，TestItems.xlsx需同程序一起修改
				程序运行过程中勿进行任何操作
	3.3 61850自动对点
		3.3.1 使用的外部软件：mmsdbg.exe；word.exe
		3.3.2 使用的外部文件：.\\dataFiles\\*.png————已经加入".\dataFiles"中
		3.3.3 运行步骤
			按顺序测试跳闸传动、自检传动、遥信传动并记录在word中
		3.3.3.4 使用说明
			输入传动每X条记录一次参数
			*.png文件需根据电脑显示分辨率重新截图
			测试开始前打开mmsDbg.exe，连接装置并切换至Reports界面
			测试开始前打开arptool.exe，连接虚拟液晶，进入跳闸传动界面。注：只进入该界面，程序自动登录用户
			点击“61850测试”btn后倒计时5s，将屏幕切换至mmsDbg.exe-->Reports界面以截屏记录测试结果，当前窗口需要切换至arptool.exe-->虚拟液晶界面以进行操作
			程序运行过程中勿进行任何操作
四、其他