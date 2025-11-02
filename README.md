# WPS ExcelTools加载项

### 1.安装依赖

```sh
npm install
```

### 2.启动项目

```sh
wpsjs debug
```

### 3.打包项目

```sh
wpsjs build
```

### 4.Mac M芯片安装

> 将7z压缩包解压到以下路径
> /Users/admin/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons

### 5.修改publish.xml

> 打开publish.xml
> 主要是修改url变量，路径是打包后的7z文件
> 注意name和version变量是否对应package.json文件中的变量

```xml
<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin install="null" name="excel-tools" enable="enable_dev" url="/Users/admin/Desktop/excel-tools.7z" type="et" version="1.0.0" customDomain=""/>
</jsplugins>
```

### 问题

#### wpsjs build无法打包成.7z文件

> 1. 首先全局安装npm依赖
>    npm install github:develar/7zip-bin
> 2. 修改依赖的权限
>    sudo chmod -R 777 依赖路径
