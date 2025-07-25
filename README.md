

<div align="center">
    <h1>Excel批量生成MobaXterm连接会话</h1>
</div>
<div align="center">
    <a href="./README.md"><img src="https://img.shields.io/badge/README-中文-red"></a>
    &nbsp;&nbsp;&nbsp;&nbsp;
    <a href="./LICENSE"><img src="https://img.shields.io/badge/license-Apache--2.0-yellow"></a>
    &nbsp;&nbsp;&nbsp;&nbsp;
</div>

<br>

**createMobaXtermSession**是一个可以通过模板文件定义好的连接数据生成MobaXterm工具可以导入的`MobaXterm_Sessions.mxtsessions`会话文件。

支持生成默认22连接和穿透连接两种方式。

## 快速使用

请到发布页面下载对应的安装包：[Release page](https://github.com/1976083684/createMobaXtermSession/releases)<br>

运行程序后执行3个步骤即可生成：

1. 输入模板文件路径
2. 输入要连接的用户名
3. 输入项目名称，即最外层文件夹名

![image-20250708153758087](README/image-20250708153758087.png)

1. 模板最后一项为空则采用远程IP的22端口连接
2. 模板最后一项非空则采用远程穿透IP、端口连接

![image-20250708153454194](README/image-20250708153454194.png)



最后将生成的`MobaXterm Sessions.mxtsessions`文件导入MobaXterm后如下：

![image-20250708154050294](README/image-20250708154050294.png)



## 许可

createMobaXtermSession是根据Apache-2.0许可证提供的 - 有关详细信息，请参阅[许可证文件](./LICENSE)。

