/*
 * pStyleArr :{
 *           "id":,  //
 *           "val":
 *       }
 * */

const style = function(xmlStr){
    
    return xmlStr.match(/<w:style w:type\=\"paragraph\"[\s\S]+?<\/w:style>/ig).map((item)=> {
        var iRe = /w:styleid\=\"([\s\S]+?)\"[\s\S]+?w:name[\s\S]+?w:val\=\"([\s\S]+?)\"[\s\S]+?/ig;
        var match = iRe.exec(item);
        return {
            "id": match[1],
            "val": match[2]
        }
    });
    
    
    
};





module.exports = style;
