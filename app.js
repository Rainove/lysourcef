var path = require("path");
var express = require("express");
const { send } = require("process");
const { json } = require("express");
var app = express();
app.get("/index", function (req, res) {
  var fs = require("fs");
  var xlsx = require("node-xlsx");

  var ajax = require("./ajax.js");
  start();
  function start() {
    ajax.ajax({
      url: "https://docs.qq.com/sheet/DQ09YdExRdFBIb25p?u=22782187e18f432d9f1ce7df31ab2602&tab=BB08J2&_t=1657693222373",
      type: "GET",
      data: {
        domainId: 300000000,
        localPadId: "dcVfECcucDOi",
        getCollected: 1,
        getIsAddedMy: 1,
        xsrf: "f5734fc8ef491a2e",
      },
      success: function (data) {
        var myDatas = [];
        var datas = JSON.parse(JSON.stringify(data)).datas;
        // console.log(JSON.parse(JSON.stringify(data)).split('<td class="C0">'))
        var datass = JSON.parse(JSON.stringify(data)).split('<td class="C0">');
        var jsonDatas = [];
        datass.forEach((value, index, obj) => {
          if (index == 0) return;
          var jsonData = {};
          var newDatas = value
            .replace(/<(style|script|iframe)[^>]*?>[\s\S]+?<\/\1\s*>/gi, "-")
            .replace(/<[^>]+?>/g, "-")
            .replace(/\s+/g, "-")
            .replace(/ /g, "-")
            .replace(/>/g, "-")
            .replace(/&nbsp;/g, "-")
            .split("--")
            .slice(1, 3);
          jsonData["编号"] = index;
          jsonData["姓名"] = newDatas[0];
          jsonData["状态"] = newDatas[1];
          jsonDatas.push(jsonData);
        });
      var flag = 0

      var name = req.query['name'];
      var jsonArrays = []
      jsonArrays.push(0)
      console.log(name)
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.setHeader('Access-Control-Allow-Headers', '*');
      jsonDatas.forEach((obj,index,arrays)=>{
        if(flag == 1)
         return;
        if(name === obj['姓名']){
          console.log(obj)
          console.log('查到了编号为:'+obj['编号']+"姓名为:"+obj["姓名"]+"状态为:"+obj["状态"])
          flag = 1
          jsonArrays[0]=1
          jsonArrays[1]=obj
          res.send(JSON.stringify(jsonArrays));
        }
        if(index == arrays.length-1){
          console.log("对不起，查不到")
          res.send(JSON.stringify(jsonArrays));
        }
      })
        // let txt = JSON.parse(JSON.stringify(data)).substring('<td></td>,', "", txt)
        // txt = txt.replace("  ", "").strip().split('<td class="C0">')
        // console.log('txt', txt);

        var count = 0;
        for (var index in datas) {
          var account = datas[index];
          var colum = [];
          var names;
          if (index == 0) {
            names = [];
          }
          for (var index2 in account) {
            if (index == 0) names.push(index2);
            var value = account[index2];
            if (value == null) {
              value = "";
            }
            colum.push(value);
            //                    console.log(account);
          }
          if (index == 0) {
            myDatas.push(names);
          }
          myDatas.push(colum);

          if (index == datas.length - 1) {
            writeXls(myDatas);
          }
        }
        // console.log(myDatas.length);
      },
    });
  }
  function writeXls(datas) {
    var buffer = xlsx.build({ worksheets: [{ name: "Group", data: datas }] });
    fs.writeFileSync("Group.csv", buffer, "binary");
  }
  function parseXls() {
    var obj = xlsx.parse("myFile.xlsx");
    console.log(obj);
  }
});
app.listen(4004, function () {
  console.log("web服务器开启了");
});