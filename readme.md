# OWA实现Word在线编辑（以GridFs存储文件）

### 介绍

OWA全称Office Web App，是用来做Office文档预览的一个很好的工具，通过一些配置文件的设置还可以实现Excel、PowerPoint的在线编辑功能，但对Word在线编辑并不支持。GitHub上有大牛实现了Word的在线编辑，不过文件必须是物理文件，本文在此基础上做了些扩展可以支持文件存储在分布式文件系统中，本文以GridFs为例。

### WopiHost

WopiHost是GitHub上大牛实现的使用OWA进行Word在线编辑的开源程序，代码地址[https://github.com/marx-yu/WopiHost.git](https://github.com/marx-yu/WopiHost.git)，该程序的代码结构如下图：
![](media/14538153573439/14538187926396.jpg)￼

使用该程序需要注意的几个地方：
1. 该程序依赖Microsoft.CobaltCore.dll ，这个dll在安装有OWA的服务器上可以找到，我已经将此dll文件放在项目中的Lib目录中；
2. 该程序是一个控制台程序，启动后会是一个http服务，并监听一个特定端口；
3. 基本只需关注Program类和CobaltServer两个类就可以了；

### 实现思路

要实现跟分布式文件系统的集成，需要三步：
1. 从分布式文件保存文件到服务器的某个目录中；
2. OWA加载该文件；
3. 修改文件，点击保存后文件需要存储到分布式文件系统中。

开始准备研究WopiHost的代码，并对代码进行修改来支持分布式文件系统，发现实现起来有些困难。原始的WopiHost使用方法如下：

```
http://192.168.16.98/we/wordeditorframe.aspx?WOPISrc=http://localhost:9111/wopi/files/1.docx&amp;access_token=111
```

上面的地址中`1.docx` 是需要编辑文件的文件名，根据该文件名在特定目录中去加载文件，如果换成分布式文件系统，`1.docx` 的地方应该是一个文件Id，如下：

```
http://192.168.16.98/we/wordeditorframe.aspx?WOPISrc=http://localhost:9111/wopi/files/56a6eda2c6a76910844dfe25&amp;access_token=111
```

换成Id后文件可以正确加载并保存，但因为GridFs不支持文件的更新操作，如果要更新一个文件需要先删除再添加，新添加的文件的FileId会发生变化，这样就会导致保存后页面加载的还是保存前的文件。

而且WopiHost对Word、Excel等不同类型的处理方式还不一样，即便上面遇到的问题能够解决，整个代码修改来也很复杂，所以需要转变思路，将上面提到的三个步骤分开来做：
1. 签出文件，该步骤将GridFs中的文件保存到服务器目录中；
2. 编辑文件，该步骤可以使用WopiHost的原始功能，代码不用做任何修改；
3. 签入文件，该步骤将目录中的文件保存到GridFs中。

因为WopiHost不支持一个文件同时多人编辑，所以在代码层面可以做一些独占式的控制，如一个用户将文件签出了，其他用户就不能签出，必须等编辑完签入后才能再签出编辑。

### OWAEditorWeb

一个简单的Asp.Net示例，代码地址：[https://github.com/oec2003/OWAEditorSample.git](https://github.com/oec2003/OWAEditorSample.git)