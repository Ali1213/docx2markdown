/*
 * numberObj :{
 *   numId:{
 *       abstractNumId: "",
 *       children:[
 *               {
 *                   ilvlID:,
 *                   numFmt:,
 *                   lvlText:,
 *                   start:
 *               }
 *           ]
 *   }
 * }
 */

const number = function (xmlStr) {
    let rtnObj = {};

    //<abstractNumId>标签内的各种属性
    let properties = xmlStr.match(/<w:abstractNum[\s\S]+?<\/w:abstractNum>/ig).map((item)=> {

        let idRe = /w:abstractNumId\=\"([\s\S]+?)\"/ig;
        let match = idRe.exec(item);

        let arr = item.match(/<w:lvl[\s\S]+?<\/w:lvl>/ig).map((otem)=> {
            let lvlRe = /w:lvl[\s\S]+?w:ilvl\=\"([\s\S]+?)\"[\s\S]+?<w:start[\s\S]+?w:val=\"([\s\S]+?)\"[\s\S]+?<w:numFmt[\s\S]+?w:val\=\"([\s\S]+?)\"\/>[\s\S]+?w:lvlText\s+?w:val=\"([\s\S]+?)\"[\s\S]+?<\/w:lvl>/ig;
            let match2 = lvlRe.exec(otem);
            return {
                "ilvlID": match2[1],
                "numFmt": match2[3],
                "lvlText": match2[4],
                "start": match2[2]
            }
        });

        return {
            abstractNumId: match[1],
            children: arr
        }
    });


    //<numId>标签解析 ，numId与abstractNumId之间的关系
    let relation = xmlStr.match(/<w:num\s+?w:numId\=\"([\s\S]+?)\"[\s\S]+?w:abstractNumId w:val="([\s\S]+?)"[\s\S]+?<\/w:num>/ig).map((item)=> {
        let relationRe = /<w:num\s+?w:numId\=\"([\s\S]+?)\"[\s\S]+?w:abstractNumId w:val="([\s\S]+?)"[\s\S]+?<\/w:num>/ig;
        let match = relationRe.exec(item);
        return {
            "numId": match[1],
            "abstractNumId": match[2]
        }
    });

    relation.forEach((item)=> {
        properties.some((otem)=> {
            if (otem.abstractNumId === item.abstractNumId) {
                rtnObj[item.numId] = otem;
                return true;
            }
        })
    })


    return rtnObj;

};


module.exports = number;
