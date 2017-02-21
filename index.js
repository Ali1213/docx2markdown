const fs = require("fs");
const AdmZip = require('adm-zip');
const XmlReader = require('xml-reader');
const path = require("path");
const head = require("./head.json");
const os = require("os");
const number = require("./lib/number");
const rels = require("./lib/rels");
const style = require("./lib/style");



var Translate = function (filePath) {

    this.zip = new AdmZip(filePath);
    this.zipEntries = this.zip.getEntries();

    this.docx = this.zip.readAsText("word/document.xml");
    // this.style = this.zip.readAsText("word/styles.xml");
    // this.numbering = this.zip.readAsText("word/numbering.xml");
    // this.rels = this.zip.readAsText("word/_rels/document.xml.rels");

    //缓存文件夹，存图片或其他
    this.tmpdir = path.join(os.tmpdir(), @.uuid());
    require("mew_util").makeDirSync(this.tmpdir);

    //存储，记录比标题序列号
    this.numberContext = {};

    //从document.xml.rels中取出引用链接的属性，将其放入数组中
    //数组是引用类型，存在bug

    this.relsObj = rels(this.zip.readAsText("word/_rels/document.xml.rels"));

    this.pStyleArr = style(this.zip.readAsText("word/styles.xml"));

    this.numberObj = number(this.zip.readAsText("word/numbering.xml"));

    //准备在之后将所有的页面数据存储在这个object里面，然后再次输出
    this.contentObj = {
        "head": {},
        "body": [],
        "footer": {},
        "comment": {}
    }
    this.content = "";
};

Translate.prototype.traverseNodes = function (nodes, fun) {
    var my = this;
    var text  = "";
    nodes.forEach((item)=> {
        var len = my[fun](item);
        if(typeof(len) == "undefined"){
            @dump(nodes);
        }
        text += len;
    });
    return text;
};

Translate.prototype.main = function (nodes) {
    var my = this;
    nodes.forEach((item)=> {
        if (item.type === "element" && item.name === "w:p") {
            my.content += my.traverseNodes(item.children, "paragraph")
            my.content += "\r\n\r\n\r\n"
        } else if (item.type === "element" && item.name === "w:tbl") {
            let tableStr = '{"colNum":';
            tableStr += my.traverseNodes(item.children, "table")
            tableStr += ']}';
            tableStr = tableStr.replace(/\]\,\]/g,"]]").replace(/\]\]\,\]/g,"]]]");
            my.content += my.tableStrParse(tableStr) + "\r\n\r\n\r\n";
        } else {
        @warn("document ast type : " + item.type + " And name:" + item.name + " not supported");
        }
    })
};

//将表格转换成HTML的形式
Translate.prototype.tableStrParse = function(str){

    let tableJson;
    try{
        tableJson = JSON.parse(str);
    }catch(e){
        @error("table: "+ str + " can't be parsed");
        return "";
    }
    let colNum = tableJson["colNum"];
    let tableArr = tableJson["table"];

    let storeColObj= {};
    let storeRowObj= {};
    let tempObj= {};
    let tableStr = "";

    tableArr.forEach((item,i)=>{
        let nowCol = 0;
        item.forEach((otem)=>{
            if(otem.length === 2){
                let colMatch = otem[0].match(/colspan\*(\d+)/i);
                let relMatch = otem[0].match(/rowspan\*([a-z]+)/i);

                if(relMatch){
                    if(relMatch[1]==="start"){
                        if(tempObj[nowCol]){
                            let x = parseInt(tempObj[nowCol].slice(tempObj[nowCol].indexOf("-")+1));
                            let y = parseInt(nowCol);
                            storeColObj[x+"-"+y]=parseInt(tempObj[nowCol]);
                        }
                        tempObj[nowCol] = 1+"-"+ i;
                    }else if(relMatch[1]==="on"){
                        tempObj[nowCol] = parseInt(tempObj[nowCol]) + 1 + tempObj[nowCol].slice(tempObj[nowCol].indexOf("-"));
                    }
                }


                if(colMatch){
                    storeRowObj[i+'-'+nowCol] = parseInt(colMatch[1]);
                    nowCol += parseInt(colMatch[1]);
                }

            }else if(otem.length ===3){
                let colMatch = otem[0].match(/colspan\*(\d+)/i);
                let relMatch = otem[1].match(/rowspan\*([a-z]+)/i);
                if(relMatch[1]==="start"){
                    if(tempObj[nowCol]){
                        let x = parseInt(tempObj[nowCol].slice(tempObj[nowCol].indexOf("-")+1));
                        let y = parseInt(nowCol);
                        storeColObj[x+"-"+y]=parseInt(tempObj[nowCol]);
                    }
                    tempObj[nowCol] = 1+"-"+ i;
                }else if(relMatch[1]==="on"){
                    tempObj[nowCol] = parseInt(tempObj[nowCol]) + 1 + tempObj[nowCol].slice(tempObj[nowCol].indexOf("-"));
                }

                storeRowObj[i+'-'+nowCol] = parseInt(colMatch[1]);
                nowCol += parseInt(colMatch[1]);

            } else {
                nowCol ++;
            }
        })
    });

    Object.keys(tempObj).forEach((item)=>{
        let x = parseInt(tempObj[item].slice(tempObj[item].indexOf("-")+1));
        let y = parseInt(item);
        storeColObj[x+"-"+y] = parseInt(tempObj[item]);
    });
        tableStr +=`<table>`;
        tableArr.forEach((item,x)=>{
            tableStr += "<tr>";
            item.forEach((otem,y)=>{
                if(otem.length>=2 && otem[otem.length-2] === "rowspan*on"){
                    return ;
                }
                tableStr += "<td";

                if(storeRowObj[x+"-"+y]){
                    tableStr += ` colspan="${storeRowObj[x+"-"+y]}"`;
                }

                if(storeColObj[x+"-"+y]){
                    tableStr += ` rowspan="${storeColObj[x+"-"+y]}"`;
                }
                tableStr += '>' + otem[otem.length-1];
                tableStr += "</td>"
            });
            tableStr += "</tr>";
        });
        tableStr +="</table>";


    return tableStr;

};


//处理段落
Translate.prototype.paragraph = function (node) {
    var my = this;
    var pObj = {};
    var pText = "";
    if (node.type === "element" && node.name === "w:pPr") {
        //pPr为段落样式；
        if (node.children) {
            pText += my.traverseNodes(node.children, "paragraphStyle");
        }
    } else if (node.type === "element" && node.name === "w:t") {
        if (node.children) {
            pText += my.traverseNodes(node.children, "paragraph");
        }
    } else if (node.type === "element" && node.name === "w:r") {
        if (node.children) {
            pText += my.traverseNodes(node.children, "paragraph");
        }
    } else if (node.type === "text" && node.name === "") {
        //文本文档
        pText += node.value;
        //下面是图形的处理方式，目前只处理图片
    } else if (node.type === "element" && node.name === "w:drawing") {
        if (node.children) {
            pText += my.traverseNodes(node.children, "draw")
        }
    } else {
    @warn("document ast type : " + node.type + " And name:" + node.name + " not supported By Paragraph");
    }
    return pText;
};

Translate.prototype.paragraphStyle = function (node) {
    var my = this;
    var pStyleText = "";
    if (node.type === "element" && node.name === "w:pStyle") {
        //pPr为段落样式；
        var id;
        try {
            id = node["attributes"]["w\:val"];
        } catch (e) {
            @error(e);
            return pStyleText;
        }
        my.pStyleArr.some((item)=> {
            if (item["id"] === id) {
                //如果是一级标题的话

                pStyleText += my.handleHeadNum(item["val"]);

                return true;

            }
        });

        if (node["children"]) {
            pStyleText += my.traverseNodes(node.children, "paragraphStyle")
        }
    } else if (node.type === "element" && node.name === "w:numPr") {

        if (node["children"]) {
            pStyleText += my.handleNumPr(node);
        }
    } else {
    @warn("document ast type : " + node.type + " And name:" + node.name + " not supported; by paragraphStyle");
        // @dump(node)
    }

    return pStyleText;

};

Translate.prototype.handleHeadNum = function (str) {
    var my = this;
    var headStr = "";
    let handleHeading = function(str){
        let re = /\d+/;
        let num = Number(str.match(re)[0]);
        let headStr = "";

        for(let i=1;i<=num;i++){
            headStr += "#";
        }

        headStr +=" ";
        return headStr;
    };


    switch (str) {
        case "heading 1":
        case "heading 2":
        case "heading 3":
        case "heading 4":
        case "heading 5":
        case "heading 6":
            headStr += handleHeading(str);
            break;
        default :
        @error("pStyle val= " + str + " is not supported")

    }

    return headStr

};

//处理数字的序列号（标题）
Translate.prototype.handleNumPr = function (node) {
    var my = this;
    var numStr = "";
        var ilvlID;
        var id;

        node["children"].forEach((item)=> {
            if (item["name"] === "w:ilvl") {
                ilvlID = item["attributes"]["w:val"];
            }
            if (item["name"] === "w:numId") {
                id = item["attributes"]["w:val"];
            }

        })

        //numId如果等于0，说明这个序号已经被取消掉了

        if (id == 0) {
            return numStr
        }


    try {
        my.numberObj[id].children.some((otem)=> {


            if (otem.ilvlID === ilvlID) {
                ilvlID = parseInt(ilvlID);
                if (!(my.numberContext[id] && my.numberContext[id][ilvlID]) ){
                    if(!Array.isArray(my.numberContext[id])){
                        // my.numberContext[id] = new Array(20).fill(0);
                        my.numberContext[id] = [];
                    }
                    my.numberContext[id][ilvlID] = {
                        "start": otem.start,
                        "current": parseInt(otem.start),
                        "format": otem.numFmt
                    }
                } else {
                    my.numberContext[id][ilvlID].current = my.numberContext[id][ilvlID].current + 1;
                    for(let i=my.numberContext[id].length-1;i>ilvlID;i--){
                        my.numberContext[id][i] = 0;
                    }
                }
                numStr += my.getHeadNumber(my.numberContext[id][ilvlID].current, my.numberContext[id][ilvlID].format, otem.lvlText);

                return true;
            }

        })


    } catch (e) {
        @error(e)
        @error(`下面标题格式暂不支持`);
        @error(node);
    }

    return numStr;
}

Translate.prototype.getHeadNumber = function (num, fmt, lvlText) {
    var my = this;
    var str = "";
    var arr = lvlText.match(/\%\d+?/g);
    var getValue = function (fmt, num) {
        switch (fmt) {
            case "decimal":
                return num;
                break;
            case "upperLetter":
                return head.upperLetter[num - 1];
                break;
            case "lowerLetter":
                return head.upperLetter[num - 1].toLowerCase();
                break;
            default :
                @error(`${fmt} style is not supported`);
                return "";
                break;
        }
    };

    if (arr.length > 1) {
        @error("暂不支持多级标题")
    };

    str += lvlText.replace(/\%\d+?/, getValue(fmt, num));
    return str+' ';

}
//处理表格
Translate.prototype.table = function (node) {
    var my = this;
    var tableStr = "";
    if (node.type === "element" && node.name === "w:tblGrid") {
        if (node.children) {
            my.columnNum = node.children.length;
            tableStr += `${node.children.length},"table":[`
        }
    } else if (node.type === "element" && node.name === "w:tr") {
        if (node.children) {
            tableStr += "[";
            tableStr += my.traverseNodes(node.children, "table");
            tableStr += "],";
        }
    } else if (node.type === "element" && node.name === "w:tc") {
        if (node.children) {
            tableStr += '[';
            tableStr += my.traverseNodes(node.children, "table");
            tableStr += '"],';
        }
    } else if (node.type === "element" && node.name === "w:tcPr") {
        if (node.children) {
            tableStr += my.traverseNodes(node.children, "table");
            tableStr += '"'
        }
    } else if (node.type === "element" && node.name === "w:vMerge") {
        if (node["attributes"]["w:val"]) {
            tableStr += '"rowspan*start",';
        } else {
            tableStr += '"rowspan*on",';
        }

    } else if (node.type === "element" && node.name === "w:gridSpan") {
        if (node["attributes"]["w:val"]) {
            tableStr += `"colspan*${node["attributes"]["w:val"]}",`
        }
    } else if (node.type === "element" && node.name === "w:p") {
        tableStr += my.traverseNodes(node.children, "paragraph");
        tableStr += "<br \>";
    } else {
    @warn("document ast type : " + node.type + " And name:" + node.name + " not supported; by table");
    }

    return tableStr;

}
//处理图形，目前只能处理图片
Translate.prototype.draw = function (node) {
    var my = this;
    var drawStr = "";
    if (node.type === "element" && node.name === "wp:inline") {
        if (node.children) {
            drawStr += my.traverseNodes(node.children, "draw");
        }
    } else if (node.type === "element" && node.name === "a:graphic") {
        if (node.children) {
            drawStr += my.traverseNodes(node.children, "draw");
        }
    } else if (node.type === "element" && node.name === "a:graphicData") {
        if (node.children) {
            drawStr += my.traverseNodes(node.children, "draw");
        }
    } else if (node.type === "element" && node.name === "pic:pic") {
        if (node.children) {
            drawStr += my.traverseNodes(node.children, "draw");
        }
    } else if (node.type === "element" && node.name === "pic:blipFill") {
        if (node.children) {
            drawStr += my.traverseNodes(node.children, "draw");
        }
    } else if (node.type === "element" && node.name === "a:blip") {
        var id;
        try {
            id = node["attributes"]["r:embed"];
        } catch (e) {
            return this.reject(e);
        }

        if (my.relsObj[id]) {
            var filename = path.basename(my.relsObj[id].target);
            var buf = my.zip.readFile(my.relsObj[id].target);
            fs.writeFile(path.join(my.tmpdir, filename), buf, (err)=> {
                if (err) {
                    @error(item.target + "未成功写入" + path.join(my.tmpdir, filename))
                }
            });
            // my.content += "![](" + path.join(my.tmpdir, filename) + ")";
            drawStr += "![](" + path.relative(__dirname,path.join(my.tmpdir, filename)) + ")";
        } else {
            @error("the value of id = " + id + +" can't not find in rels files");
        }

    } else {
     @warn("document ast type : " + node.type + " And name:" + node.name + " not supported; by draw");
    }

    return drawStr;
};

Translate.prototype.render = function () {
    var my = this;
    return @.async(function () {
        const reader = XmlReader.create();
        reader.on('done', data => this.next(data));
        reader.parse(my.docx);
    }).then(function (result) {
        if (result.name === "w:document" && result.type === "element") {
            if (result.children[0].type === "element" && result.children[0].name === "w:body") {
                my.main(result.children[0].children);
            } else {
                @error("can't find w:body label, please checkout!")
            }
        } else {
            @error("can't find w:document label, please checkout!");
        }
        this.next();
    }).then(function () {
        this.next(my.content);
    });
};


module.exports = function (path) {
    var turn = new Translate(path);
    return turn.render();
}