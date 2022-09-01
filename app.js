const xlsx = require("node-xlsx");
const puppeteer = require("puppeteer");
const express = require("express"); //引入express 模块
const app = express(); //创建实例
const path = require("path");
app.use("/", express.static(path.join(__dirname, "./")));
const fs = require("fs"); // 引入文件系统模块
var syncRequest = require("sync-request");
var CronJob = require("cron").CronJob;
const { appConfig } = require("./config.js");
const {
  promises: { readFile, writeFile },
} = require("fs");
const { segment } = require("oicq");
const { createClient, Message } = require("oicq");
//服务器首页
app.get("/", function (req, res) {
  res.sendFile(__dirname + "/" + "index.html");
});

//服务器获取excel数据
app.get("/excel", async function (req, res) {
  let data = await getExcel(req.query.time, req.query.grade);
  res.send(data);
});

// 开启服务器
app.listen(2999, () => {
  console.log("服务器在localhost:2999开启。。。。。");
});
//获取截图
async function getPics(grade) {
  let time = new Date().getTime();

  await getXls({ password: appConfig.password, sn: appConfig.user }, time); //输入账号和密码
  await openUrl(time, grade);
  return time;
}

//获取未打卡的Excel
async function getXls(data, fileName) {
  var res = syncRequest(
    "POST",
    "https://yjsxx.whut.edu.cn/yqtj/api/login/loginIn",
    { json: data }
  );
  let cookie = res.headers["set-cookie"][0];
  var res = syncRequest(
    "POST",
    "https://yjsxx.whut.edu.cn/yqtj/exportDayRegisterStatistics/",
    {
      headers: { Cookie: cookie },
      json: {
        collegeName: null,
        majorName: null,
        gradeName: null,
        className: null,
        endTime: "",
        startTime: null,
        today: true,
        stuNameOrSn: null,
        isReg: 0,
        page: 1,
        size: 20,
      },
    }
  );
  await writeFile(fileName + ".xls", res.body);
}

//解析Excel数据
async function getExcel(time, grade) {
 
  var obj = xlsx.parse(time + ".xls");
  var Data = obj[0].data;
  console.log("解析excel成功");
  if (grade == "all") {
    return {
      2021: getData("2021", Data),
      2020: getData("2020", Data),
      2019: getData("2019", Data),
      2018: getData("17and18", Data), //"17and18"
      2017: getData("under17", Data), //"under17"
    };
  } else {
    return {
      [grade]: getData(grade, Data),
    };
  }
}

//筛选对应年级数据
function getData(type, data) {
  if (type == "all") {
    return data;
  }
  var temp = [];
  temp.push(data[0]); //表头
  if (type >= 2021) {
    for (i in data) {
      if (data[i][6] == type) {
        temp.push(data[i]);
      }
    }
  }
  if (type == "under17") {
    for (i in data) {
      if (data[i][5] == "马区" && data[i][6] <= 2017) {
        temp.push(data[i]);
      }
    }
  }
  if (type == "17and18") {
    for (i in data) {
      if (data[i][5] == "马区" && data[i][6] >= 2017 && data[i][6] <= 2018) {
        temp.push(data[i]);
      }
    }
  }
  for (i in data) {
    if (data[i][5] == "马区" && data[i][6] == type) {
      temp.push(data[i]);
    }
  }
  return temp;
}

//获取所有未打卡名字
async function getName(time, grade) {
  var obj = xlsx.parse(time + ".xls");
  var Data = getData(grade, obj[0].data);
  let names = [];
  for (let item of Data) {
    names.push(item[1]);
  }
  console.log(names.join(","));
  return names;
}

//开启浏览器下载
async function openUrl(time, grade) {
  const browser = await puppeteer.launch({
    headless: true,
    defaultViewport: { width: 700, height: 986 },
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  const page = await browser.newPage(); //打开一个新页面
  await page.goto("http://127.0.0.1:2999?time=" + time + "&grade=" + grade);
  //设置下载路径
  const client = await page.target().createCDPSession();
  client.send("Page.setDownloadBehavior", {
    behavior: "allow",
    downloadPath: __dirname,
  });
  //等待图片下载完成
  if (grade == "all") {
    for (let i = 2017; i <= 2021; i++) {
      let res = await waitFile(`${__dirname}/${time + "table" + i}.png`);
      console.log(res);
    }
  } else {
    let res = await waitFile(`${__dirname}/${time + "table" + grade}.png`);
    console.log(res);
  }
  await browser.close();
}

//等待文件出现
function waitFile(fileName) {
  return new Promise(function (resolve, reject) {
    let timeout = 10000; //查找十秒
    let timer;
    let timeOut = setTimeout(() => {
      clearInterval(timer);
      resolve(`文件 ${fileName} 未找到.`);
    }, timeout);
    timer = setInterval(() => {
      console.log("查找中");
      if (fs.existsSync(fileName)) {
        clearTimeout(timeOut);
        clearInterval(timer);
        resolve(`文件 ${fileName} 已出现.`);
      }
    }, 500); //每0.5s扫描一次
  });
}
//qq发送信息
async function sendMessage(group, grade) {
  let time = await getPics(grade);
  let names = await getName(time, grade);
  let memberMap = await group.getMemberMap(); //获取群成员
  let qqList = [];
  //找到群成员中未打卡的人
  memberMap.forEach(function (item) {
    for (let name of names) {
      if (item.card.match(name)) {
        qqList.push(item.user_id);
      }
    }
  });
  if (grade == "all") {
    for (let i = 2017; i <= 2021; i++) {
      await group.sendMsg(segment.image(`${time + "table" + i}.png`));
    }
  } else {
    await group.sendMsg(segment.image(`${time + "table" + grade}.png`));
  }
  let message = ["请健康打卡！！！"];
  for (let item of qqList) {
    message.push(segment.at(item)); //@对应同学
  }
  group.sendMsg(message);
}

async function qqBot() {
  const account = appConfig.qq; //登录的QQ号
  const bot = createClient(account);
  bot
    .on("system.login.qrcode", function (e) {
      this.logger.mark("扫码后按Enter完成登录");
      process.stdin.once("data", () => {
        this.login();
      });
    })
    .login();
  //开启qq机器人;
  bot.on("message.group", async function (msg) {
    console.log(msg);
    if (msg.raw_message.split("，")[0] === "未打卡") {
      await msg.group.sendMsg("正在查询...");
      let grade = msg.raw_message.split("，")[1];
      grade = grade ? grade : "all";
      await sendMessage(msg.group, grade);
    }
    if (msg.raw_message.split("，")[0] === "定时打卡") {
      let grade = msg.raw_message.split("，")[1];
      let time = msg.raw_message.split("，")[2].split("：");
      grade = grade ? grade : "all";
      runEveryDay(time[1] + " " + time[0] + " * * *", msg.group, grade);

      msg.group.sendMsg(
        "定时成功,每天" + time[0] + ":" + time[1] + "自动催打卡"
      );
    }
    if (msg.raw_message === "憨憨") {
      msg.group.sendMsg(["你才是憨憨", segment.face(104)]);
      // 发送一个戳一戳
      msg.member.poke();
    }
  });
}
qqBot();

//每天定时运行
function runEveryDay(timestr, group, grade) {
  new CronJob(
    timestr,
    function () {
      sendMessage(group, grade);
    },
    null,
    true
  );
}
//每年级为一个sheet导出为一个Excel 有需自取
async function exportExcel() {
  var obj = xlsx.parse("每日填报填报统计数据.xls");
  var Data = obj[0].data;
  //设置每一行的宽度
  const options = {
    "!cols": [
      { wpx: 100 },
      { wpx: 100 },
      { wpx: 270 },
      { wpx: 220 },
      { wpx: 120 },
      { wpx: 120 },
      { wpx: 120 },
      { wpx: 120 },
      { wpx: 180 },
    ],

    "!margins": {
      left: 0.7,
      right: 0.7,
      top: 0.7,
      bottom: 0.7,
      header: 0.3,
      footer: 0.3,
    },
  };
  var buffer = xlsx.build(
    [
      { name: "2021", data: getData("2021", Data) },
      { name: "2020", data: getData("2020", Data) },
      { name: "2019", data: getData("2019", Data) },
      { name: "2018和2017", data: getData("17and18", Data) },
      { name: "2017以下", data: getData("under17", Data) },
    ],
    options
  );
  await writeFile("导出.xls", buffer);
  console.log("导出excel成功");
}
//接受报错信息
process.on("unhandledRejection", (reason, promise) => {
  console.log("Unhandled Rejection at:", promise, "reason:", reason);
});
