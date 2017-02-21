/*
 * relsObj :{
 *       idValue:{   //idValue :id值
 *           "type":str,//文件类型
 *           "target":str//文件路径
 *           }
 *   }
 **/
const rels = function(xmlStr){

    var rtnObj = {};

    xmlStr.match(/<Relationship\s+?([\s\S]+?\/>)/ig).forEach((item) => {
        var iRe = /id\=\"([\s\S]+?)\"[\s\S]+?type\=\"([\s\S]+?)\"[\s\S]+?target\=\"([\s\S]+?)\"/ig;
        var match = iRe.exec(item);

        rtnObj[match[1]] = {
            "type": match[2],
            "target":  "word/" + match[3]
        }
    })

    return rtnObj;
};


module.exports = rels;
