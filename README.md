# 列车时刻表

一个基于 VB6 开发的 Windows 桌面端列车时刻表查询工具，通过调用 12306 的接口获取实时列车数据。

## 功能特性

- **车次查询**：输入车次号，查询该车次完整的经停站时刻表
- **车站查询**：选择车站和日期，查询该站当日所有经停列车信息
- **区间查询**：选择出发站和到达站，查询两站之间的所有列车
- **智能搜索**：支持拼音、拼音首字母、汉字等多种方式搜索车站和车次
- **异步请求**：基于 MSXML 事件驱动，UI 不阻塞，响应流畅

## 技术栈

- **语言**：Visual Basic 6.0
- **UI 控件**：[VBCCR18 (Krool's Common Controls)](https://github.com/Kr00l/VBCCR18) 1.2
- **HTTP 请求**：MSXML2.ServerXMLHTTP.6.0（异步回调模式）
- **JSON 解析**：自研纯 VB6 实现的 JSON 解析器
- **数据来源**：12306.cn

## 项目结构

```
列车时刻表/
├── 列车时刻表.vbp          # VB6 项目文件
├── 列车时刻表.vbw          # VB6 工作区文件
├── style.res               # 视觉样式资源（启用 Windows 主题）
├── FrmStart.frm            # 启动界面 - 功能菜单导航
├── FrmTrain.frm            # 车次查询窗体
├── FrmStation.frm          # 车站查询窗体
├── FrmTicket.frm           # 区间查询窗体
├── StationPicker.ctl       # 车站选择用户控件（输入联想）
├── TrainNoPicker.ctl       # 车次选择用户控件（输入联想）
├── Utilities.bas           # 通用工具模块（URL 编码等）
├── StationData.bas         # 车站数据模块（解析与搜索）
├── AsyncRequest.cls        # 异步 HTTP 请求类
├── JSON.cls                # JSON 解析类
└── ocx/
    └── VBCCR18.OCX         # Krool 通用控件库
```

## 构建与运行

### 环境要求

- Windows 操作系统
- Visual Basic 6.0 开发环境
- VBCCR18.OCX 控件（已内置在 `ocx/` 目录下）

### 注册控件

首次运行前需要注册 VBCCR18.OCX 控件：

```cmd
regsvr32 ocx\VBCCR18.OCX
```

### 编译

在 VB6 IDE 中打开 `列车时刻表.vbp`，点击「文件 → 生成 列车时刻表.exe」即可编译。

## 接口说明

项目共使用了以下 5 个 12306 接口：

### 1. 车站列表接口

获取全国所有车站的名称、电报码、拼音等信息，用于车站搜索联想。

- **URL**：`GET https://kyfw.12306.cn/otn/personalJS/core/common/station_name_new.js`
- **返回格式**：`var station_names = '@bjb|北京北|VAP|beijingbei|bjb|...'`

### 2. 车次搜索接口

根据关键词搜索车次，用于车次输入联想以及车次查询的第一阶段。

- **URL**：`GET https://search.12306.cn/search/v1/train/search?keyword={keyword}&date={yyyyMMdd}`
- **返回**：JSON，包含 `data` 数组，每项含 `train_no`、`station_train_code`、`from_station`、`to_station` 等字段

### 3. 车次停靠站查询接口

根据车次内部编号查询完整经停站时刻表，用于车次查询的第二阶段。

- **URL**：`GET https://kyfw.12306.cn/otn/queryTrainInfo/query?leftTicketDTO.train_no={train_no}&leftTicketDTO.train_date={yyyy-MM-dd}`
- **返回**：JSON，`data.data` 数组包含每站的 `station_no`、`station_name`、`arrive_time`、`start_time`、`running_time` 等字段

### 4. 区间票价查询接口

根据出发站、到达站、日期查询两站之间的所有列车及票价信息。

- **URL**：`GET https://kyfw.12306.cn/otn/leftTicketPrice/queryAllPublicPrice?leftTicketDTO.train_date={yyyy-MM-dd}&leftTicketDTO.from_station={code}&leftTicketDTO.to_station={code}&purpose_codes=ADULT`
- **返回**：JSON，`data` 数组包含 `queryLeftNewDTO` 对象，其中含 `station_train_code`、`train_class_name`、`lishi`、`swz_price`、`zy_price`、`ze_price` 等字段

### 5. 车站大屏接口

查询指定车站在指定日期的所有经停列车信息（微信小程序接口）。

- **URL**：`POST https://mobile.12306.cn/wxxcx/wechat/bigScreen/queryTrainByStation`
- **Content-Type**：`application/x-www-form-urlencoded`
- **参数**：
  - `train_station_code`：车站电报码（如 BJP 表示北京）
  - `train_start_date`：查询日期（格式 YYYY-MM-DD，如 2026-08-08）
- **请求头**：需模拟微信小程序环境
  - `Referer`：`https://servicewechat.com/`
- **返回**：JSON，`data` 数组包含 `station_train_code`、`train_class_name`、`start_time`、`arrive_time`、`platform_no`、`jiaolu_corporation_code` 等丰富字段

## 关于

- **作者**：Cr4zy Uru5
- **网站**：[crazyurus.com](https://crazyurus.com/)
- **版本**：1.0.13