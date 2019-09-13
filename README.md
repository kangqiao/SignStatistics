
# 签到信息统计工具

为了方便[仁爱](chrenai.org)心栈项目与[志愿北京](www.bv2008.cn)对接, 开发此工具, 旨在将仁爱活动签到信息收集整理并统计为Excel表格数据, 方便录入志愿北京系统中.

## 环境搭建

### Python
- 安装 [python](https://realpython.com/installing-python/)
- 安装 [virtualenv](https://virtualenv.pypa.io/en/latest/installation/)
- 安装 [virtualenvwrapper](https://virtualenvwrapper.readthedocs.io/en/latest/install.html)

### Install (Optional)
通过virtualenv安装依赖包
```bash
~ git clone git@github.com:kangqiao/SignStatistics.git
~ cd SignStatistics
~ mkvirtualenv --no-site-packages --python=python3.7 SignStatistics
~ pip install -r requirements.ini
```
### 配置
#### 名字映射配置
主要解决签到时, 由于签错字导致的同一人多个名字, 在不更改原签到信息的情况下, 需要配置签错的名字为别名.
```ini
# 格式如下:
# ID=真实姓名,别名1,别名2,别名3
11010000000001=张三,zhangsan
11010000000002=韩梅梅,梅梅
```
#### 工时配置
心栈每天活动从熬粥到奉粥结束, 总工时需要4小时. 
- 熬粥志愿者由于4点就来了, 所以会有2小时熬粥工时, 但最多4个小时.
- 普通志愿者若只是奉粥, 则只有1小时工时.
- 普通志愿者除参与奉粥外, 也积极做些后勤工作, 会相应增加部分工时, 但不会超过2小时.

在`model.py`中的Hour中可配置不同志愿活动所需的工时.
```python
# 熬粥2 日负责2 文宣2 后勤0.5 奉粥1 环保1
Hour = {
    SIGN_COOK_GRUEL: 3,
    SIGN_MANAGER: 2,
    SIGN_PUBLICITY: 3,
    SIGN_SUPPORTER: 0.5,
    SIGN_SERVICE: 1,
    SIGN_PROTECT_ENV: 1
}
```

### Run
```python
~ python main.py path='.' output=X月份签到统计表
```
> path参数指定签到文件路径, 或者签到文件所在目录(目录下可以存在多个签到文件) , 默认为当前路径`'.'`

> output参数指定最后输出的Excel文件名, 默认为`签到统计表`

### 签到信息文件内容格式
签到信息文件扩展名必须为`.txt`, 格式如下:
```text
标题：XXX-xxx
奉粥日期：2019年5月24日（周五）
日负责人：AAA、BBB
签到：AAA
熬粥：CCC
前行：BBB、CCC、DDD
杯数：287 杯
人数：25 人
新人数：1 人，FFF
摄影：AAA
日志：
文宣：EEE
结行：DDD
后勤：EEE、FFF
环保：FFF
奉粥：AAA、BBB、CCC、DDD、EEE、FFF

标题：XXX-xxx
奉粥日期：2019年5月25日（周六）
日负责人：AAA、BBB
签到：AAA
熬粥：CCC
前行：BBB、CCC、DDD
杯数：287 杯
人数：25 人
新人数：1 人，FFF
摄影：AAA
日志：
文宣：EEE
结行：DDD
后勤：EEE、FFF
环保：FFF
奉粥：AAA、BBB、CCC、DDD、EEE、FFF

......
```
> 注: 每天签到信息间以空行隔开, 每行的标题与内容间以冒号隔开(`r'[：|:]'`), 人名之间以顿号,逗号,空格(`r'[、|,|，|\s]'`)都可以,



**仁爱无限, 善愿承办**
```
////////////////////////////////////////////////////////////////////
//                          _ooOoo_                               //
//                         o8888888o                              //
//                         88" . "88                              //
//                         (| ^_^ |)                              //
//                         O\  =  /O                              //
//                      ____/`---'\____                           //
//                    .'  \\|     |//  `.                         //
//                   /  \\|||  :  |||//  \                        //
//                  /  _||||| -:- |||||-  \                       //
//                  |   | \\\  -  /// |   |                       //
//                  | \_|  ''\---/''  |   |                       //
//                  \  .-\__  `-`  ___/-. /                       //
//                ___`. .'  /--.--\  `. . ___                     //
//              ."" '<  `.___\_<|>_/___.'  >'"".                  //
//            | | :  `- \`.;`\ _ /`;.`/ - ` : | |                 //
//            \  \ `-.   \_ __\ /__ _/   .-` /  /                 //
//      ========`-.____`-.___\_____/___.-`____.-'========         //
//                           `=---='                              //
//      ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^        //
//                          南无阿弥陀佛                           //
////////////////////////////////////////////////////////////////////

```