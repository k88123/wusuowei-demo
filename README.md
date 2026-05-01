1. MVVM / 数据绑定怎么实现
用 ObservableCollection<T> 绑定 DataGrid，自动刷新列表
数据模型实现 INotifyPropertyChanged，单元格数值实时更新
不重新 new 集合，只用 Clear() + Add() 保证绑定不失效
2. USB HID 通信逻辑
一报换一报：OnReportReceived 收到后立刻再 ReadReport
构建异步监听环路，不丢数据、不阻塞
字节解析：(data[1] << 8) | data[0] 拼 16 位 ADC 值
3. 0.1℃ 温变触发采样算法
防抖：Math.Abs(当前温 - 上次记录) ≥ 0.1 才记录
自动识别升温 / 降温方向
只存有效数据，不存噪声
4. 多线程 & 异步 & 界面不卡死
async/await 处理 USB 指令，UI 不卡顿
Dispatcher 安全跨线程更新 UI
Invoke（同步等待） / BeginInvoke（异步不等待） 区分使用场景
狂点按钮不会崩：按钮禁用 + isProcessing 拦截 + CancellationToken
5. 线程选型
Task.Run：USB 通信、耗时操作
Thread：长驻后台监听
Parallel.ForEach：大数据计算
DispatcherTimer：定时刷新 UI
