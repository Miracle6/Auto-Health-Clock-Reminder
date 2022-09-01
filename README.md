# Auto-Health-Clock-Reminder
使用Node.JS开发的自动QQ催同学健康打卡；

武汉理工大学完美适用（需要有后台管理的账号！！！）

其余学校的同学可以借鉴思想对代码进行修改。

### 可解决问题

- 可以自动获取学院每天未打卡的同学Excel，并可以对Excel数据分年级进行解析。

- 将学院每天未打卡的同学数据生成图片
- QQ自动催同学打卡，发送未打卡图片，@未打卡的同学

### 使用说明

1. 建议在服务器上搭建，可以一劳永逸。

2. 安装最新版Node

   Win直接去官网 [Node.js (nodejs.org)](https://nodejs.org/zh-cn/)

   linux可以参考这个[[centos上安装forever - IT教主 - 博客园 (cnblogs.com)](https://www.cnblogs.com/tengqf/articles/14234093.html)](https://blog.csdn.net/weixin_42678675/article/details/124033131)

3. 下载本代码文件夹

4. `cmd`切换到文件夹根目录，运行`npm i` 安装程序。要是有报错，使用`cnpm`来安装

   ```source-shell
   npm install -g cnpm --registry=https://registry.npm.taobao.org
   cnpm i
   ```

5. 去`config.js`里面配置您的账号信息

6. QQ登录默认使用扫码登录（可能不太稳定）

   有时间折腾可以使用如下方法：

   [01.使用密码登录 (滑动验证码教程) · takayama-lily/oicq Wiki (github.com)](https://github.com/takayama-lily/oicq/wiki/01.使用密码登录-(滑动验证码教程))

7. `node app.js`即可以开启服务，Linux可以安装forever模块，使永远在服务器上运行。

8. Enjoy it！

### Tips

设置定时格式，可以参考[Crontab.guru - The cron schedule expression editor](https://crontab.guru/#00_15_*_*_*)