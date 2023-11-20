## 需求

我需要开发一款软件 我打算使用python3来做 gui用PyQt5 这个软件的作用是通过websocket和客户端进行双向通信 在客户端发出某些指令的情况下 将客户端提供的json数据写进excel 读写excel使用xlwings库 这个库可以在不关闭excel表格的情况下读写excel很方便 我来描述一下窗口如何设计
一共有三个输入框 一个多行输入框 
两个多选框 以上都是输入控件都用于保存配置信息 类似表单功能 这是以上六个控件的数据伪代码 {
    boss_name: 'unknown',
    max_times: '9999',
    check_timeout: '6',
    friend: '',
    is_multiple: true,
    is_draw: true,
}
接下来我需要三个按钮 分别是清空配置、保存配置、开始运行。
清空配置按钮很简单 就是清空以上六个输入控件的信息 保存配置则将以上六个输入控件的表单数据通过websocket服务发送到客户端 格式为
{
    code: 1,
    call: "save_config",
    data:{
        boss_name: 'unknown',
        max_times: '9999',
        check_timeout: '6',
        friend: '',
        is_multiple: true,
        is_draw: true,
    }
}
接下来我给你说软件的大概启动流程
1. 启动qt5窗口
2. 创建websocket服务
3. 监听客户端消息 该消息为json数据 格式为{
    code: 1,
    call: "",
    data: "",
} code用于区分客户端发送的消息类型 消息类型有两种 用1和2区分 分别为单向消息 和双向消息 单向消息就是收到客户端消息后不用回复 双向消息就是收到客户端消息后 需要返回一些数据; call参数为客户端传入的函数名 比如登录的时候客户端传入login 此时我判断如果call=="login" 就知道客户端登录了 data为客户端传入的参数数据 有可能为空
客户端首次登录的时候 会发送两条json数据过来 第一条是
{
    code: 1,
    call: "login",
    data: ""
}
这条不用处理 或者简单的打印一个print消息就行
第二条是
{
    code: 1,
    call: "set_config",
    data:{
        boss_name: 'unknown',
        max_times: '9999',
        check_timeout: '6',
        friend: '',
        is_multiple: true,
        is_draw: true,
    }
}
如何收到这条消息 就得更新软件的六个输入框信息
最后 客户端会不间断发送一条json信息 你要将其写入excel 例子如下
{ header: 
    { golds: 0,
      win: 0,
      lose: 0,
      draw: 0,
      cattle: 9,
      stake_count: 0,
      stake_golds: 0,
      current_golds: 0,
      max_times: 9999,
      name: 'chaod',
      redp: '0.45',
      current_count: 1,
      wld: '0/0/0',
      key: '0_1700326000922' },
   body: 
    [ { name: 'Chaos',
        golds: 0,
        current_golds: 0,
        chat: 0,
        cattle: 3,
        redp: '0.21',
        isfake: true,
        islazy: false,
        isbug: true },
      { name: '；PianJ',
        golds: 0,
        current_golds: 0,
        chat: 0,
        cattle: 7,
        redp: '0.34',
        isfake: true,
        islazy: false,
        isbug: true },
      { name: '墨渊战神',
        golds: 0,
        current_golds: 0,
        chat: 1,
        cattle: 0,
        redp: '0.00',
        isfake: true,
        islazy: false,
        isbug: true },
      { name: '碧水蓝天',
        golds: 0,
        current_golds: 0,
        chat: 13,
        cattle: 0,
        redp: '0.00',
        isfake: true,
        islazy: false,
        isbug: true }
    ]
}
header为头部信息 在表格顶部展示 如果有新的header数据收到 就覆盖写入excel顶部区域
body数据类型是一个数组 为行数据 每一行显示body中的一个元素 如果有新的body数据收到 不要覆盖原来的数据 而是继续往下面写入 记得两次数据直接隔一个空行数据 



{
    code: 0,
    call: "",
    data: "",
}

uidata = {
    boss_name: 'unknown',
    max_times: '9999',
    check_timeout: '6',
    friend: '',
    is_multiple: true,
    is_draw: true,
}



{'code': 1, 'call': 'getGameDataCurrent', 'data': {'header': {'golds': 0, 'win': 0, 'lose': 0, 'draw': 0, 'cattle': 9, 'stake_count': 0, 'stake_golds': 0, 'current_golds': 0, 'max_times': 9999, 'name': 'chaod', 'redp': '0.45', 'current_count': 2}, 'body': [{'name': 'Chaos', 'golds': 0, 'current_golds': 0, 'chat': 0, 'cattle': 3, 'redp': '0.21', 'isfake': True, 'islazy': False, 'isbug': True}, {'name': '；PianJ', 'golds': 0, 'current_golds': 0, 'chat': 0, 'cattle': 7, 'redp': '0.34', 'isfake': True, 'islazy': False, 'isbug': True}, {'name': '墨渊战神', 'golds': 0, 'current_golds': 0, 'chat': 1, 'cattle': 0, 'redp': '0.00', 'isfake': True, 'islazy': False, 'isbug': True}, {'name': '碧水蓝天', 'golds': 0, 'current_golds': 0, 'chat': 13, 'cattle': 0, 'redp': '0.00', 'isfake': True, 'islazy': False, 'isbug': True}]}}