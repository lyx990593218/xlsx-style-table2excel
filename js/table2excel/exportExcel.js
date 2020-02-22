/*
* 基于xlsx js封装的导出layui table成excel
* @createtime 2020/2/12 今天很下雨，这个js也很冷清
* @author Douglas Lai
*/

var exportExcel = (function (window) {
    var exportExcel = function () {
        return new exportExcel.fn.init();
    };
    // content begin
    exportExcel.fn = exportExcel.prototype = {
        templateProps: {},
        lay_id: null,
        configTableBefore: function () {

        },
        configTableAfter: function () {

        },
        filename: null,
        renderSheet: function () {

        },
        hpxStartRowsIndex: 0,
        wchStartRowsIndex: 1,

        headerColSpanNum : 0,
        hpx : [],
        wch : [],
        arr : ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"],
        sheet: null,
        tempTrArr : null,

        totalWch : 10 * 20,

        // 生命周期
        init: function () {
            layer = top.layer || layer;// 兼容框架内layer
            this.yinYongJS();
        },
        reset: function () {
            // 事件解绑
            if (document.addEventListener) {
                document.addEventListener('DOMMouseScroll', function () {
                }, false);
            }//W3C
            window.onmousewheel = document.onmousewheel = function () {
            };//IE/Opera/Chrome

            window.onmousemove = null;
            window.onmouseup = null;
        },

        setProps: function(props){
            var that = this;

            that.templateProps = props;

            that.lay_id = props.lay_id;

            if (props.configTableBefore){
                that.configTableBefore = props.configTableBefore;
            }

            if (props.configTableAfter){
                that.configTableAfter = props.configTableAfter;
            }

            that.filename = props.filename;

            if (props.renderSheet){
                that.renderSheet = props.renderSheet;
            }

            if (props.hpxStartRowsIndex){
                that.hpxStartRowsIndex = props.hpxStartRowsIndex;
            }

            if (props.wchStartRowsIndex){
                that.wchStartRowsIndex = props.wchStartRowsIndex;
            }

            if (props.totalWch){
                that.totalWch = props.totalWch;
            }
        },
        /**
         *
         * @param props.lay_id table id
         * @param props.configTable 配置临时表，可动态插入表头
         * @param props.filename 文件名
         * @param props.renderSheet 渲染sheet页，设置宽、高、居中、边框线等
         */
        tableToExcel: function (props) {
            var that = this;

            that.createExportTable();

            that.setProps(props);

            // 生成临时表
            that.newTempTable();

            that.sheet = XLSX2.utils.table_to_sheet($('#exportTable')[0], {rowIndex: 0});

            that.tempTrArr = $('#exportTable').find('tr'); // 临时表的所有tr

            that.configHpx(props.hpxStartRowsIndex);
            that.configWch(props.wchStartRowsIndex);

            props.renderSheet(that.arr, that.hpx, that.wch, that.sheet, that.headerColSpanNum);

            that.sheet["!cols"] = that.wch;
            that.sheet['!rows'] = that.hpx;
            that.openDownloadDialog(that.sheet2blob(that.sheet, props.filename), props.filename + '.xlsx');
        },

        // 配置高度
        configHpx : function (startRowsIndex) {
            var that = this;

            var tempTrArr = that.tempTrArr;
            var arr = that.arr;
            var sheet = that.sheet;

            for (let i = that.hpxStartRowsIndex; i < tempTrArr.length; i++) {  // 从表格的第四行开始
                for (let j = 0; j < arr.length; j++) {
                    if (sheet[arr[j] + Number(i + 1)]) {
                        sheet[arr[j] + Number(i + 1)].s = {
                            font: {
                                name: '宋体',
                                sz: 12,
                                bold: false,
                                underline: false,
                                /*color: {
                                    rgb: "FFFFAA00"
                                }*/
                            },
                            alignment: {horizontal: "center", vertical: "center", wrapText: true},
                            border: {
                                top: {style: 'thin'},
                                left: {style: 'thin'},
                                bottom: {style: 'thin'},
                                right: {style: 'thin'}
                            }
                            /*fill: {
                                bgColor: { rgb: 'ffff00' }
                            }*/
                        };
                    } else {
                        sheet[arr[j] + Number(i + 1)] = {
                            s: {
                                border: {
                                    top: {style: 'thin'},
                                    left: {style: 'thin'},
                                    bottom: {style: 'thin'},
                                    right: {style: 'thin'}
                                }
                            }
                        };
                    }
                }

                if (that.hpx[i]) continue;
                that.hpx[i] = {hpx: 35};
            }
        },

        //配置宽度
        configWch : function () {
            var that = this;

            var tempTrArr = that.tempTrArr;

            var headerTh = tempTrArr.eq(that.wchStartRowsIndex).find("td");

            for (var thIndex = 0; thIndex < headerTh.length; thIndex++) {
                that.wch[thIndex] = {wch: Math.floor(that.totalWch / headerTh.length)}; // 第二列开始
            }
        },

        newTempTable : function () {
            var that = this;

            var hradertrArr = $("div[lay-id='"+that.lay_id+"'] .layui-table-header>.layui-table").find("tr"); //所有行
            var trArr = $("div[lay-id='"+that.lay_id+"'] .layui-table-main>.layui-table").find("tr"); //所有行

            // 真实数据行td数量，用于设置表头所需， 配置宽高等
            that.headerColSpanNum = (trArr.eq(0).find('td').length);

            that.configTableBefore();

            that.appTr2Table(hradertrArr);
            that.appTr2Table(trArr);

            that.configTableAfter();
        },

        appTr2Table : function (trArr) {
            for (var trIndex = 0; trIndex < trArr.length; trIndex++) {
                var curTr = trArr.eq(trIndex);
                // var curTrFirstTd = curTr.find('td').eq(0);

                $('#exportTable').append(curTr[0].outerHTML);
            }
        },

        createExportTable : function (){
            $('#exportTable').remove();
            //创建一个table
            var table = document.createElement('table');
            table.setAttribute("id","exportTable");
            table.setAttribute("style","display: none;");

            var bo = document.body; //获取body对象.
            //动态插入到body中
            bo.insertBefore(table, bo.lastChild);
        },

        yinYongJS : function(){
            var new_element=document.createElement("script");
            new_element.setAttribute("type","text/javascript");
            new_element.setAttribute("src","js/table2excel/xlsx.extendscript.js");
            document.body.appendChild(new_element);

            var new_element2=document.createElement("script");
            new_element2.setAttribute("type","text/javascript");
            new_element2.setAttribute("src","js/table2excel/xlsx-style/xlsx.full.min.js");
            document.body.appendChild(new_element2);
        },

        // 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
        sheet2blob : function(sheet, sheetName) {
            sheetName = sheetName || 'sheet1';
            var workbook = {
                SheetNames: [sheetName],
                Sheets: {}
            };
            workbook.Sheets[sheetName] = sheet;
            // 生成excel的配置项
            var wopts = {
                bookType: 'xlsx', // 要生成的文件类型
                bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
                type: 'binary'
            };
            var wbout = XLSX.write(workbook, wopts);
            var blob = new Blob([s2ab(wbout)], {type: "application/octet-stream"});

            // 字符串转ArrayBuffer
            function s2ab(s) {
                var buf = new ArrayBuffer(s.length);
                var view = new Uint8Array(buf);
                for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                return buf;
            }

            return blob;
        },

        /**
         * 通用的打开下载对话框方法，没有测试过具体兼容性
         * @param url 下载地址，也可以是一个blob对象，必选
         * @param saveName 保存文件名，可选
         */
        openDownloadDialog : function (url, saveName) {
            if (typeof url == 'object' && url instanceof Blob) {
                url = URL.createObjectURL(url); // 创建blob地址
            }
            var aLink = document.createElement('a');
            aLink.href = url;
            aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
            var event;
            if (window.MouseEvent) event = new MouseEvent('click');
            else {
                event = document.createEvent('MouseEvents');
                event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
            }
            aLink.dispatchEvent(event);
        },

        // 日期格式化
        dateFormat: function (fmt, date) {
            var ret;
            var opt = {
                "Y+": date.getFullYear().toString(),        // 年
                "m+": (date.getMonth() + 1).toString(),     // 月
                "d+": date.getDate().toString(),            // 日
                "H+": date.getHours().toString(),           // 时
                "M+": date.getMinutes().toString(),         // 分
                "S+": date.getSeconds().toString()          // 秒
                // 有其他格式化字符需求可以继续添加，必须转化成字符串
            };
            for (var k in opt) {
                ret = new RegExp("(" + k + ")").exec(fmt);
                if (ret) {
                    fmt = fmt.replace(ret[1], (ret[1].length == 1) ? (opt[k]) : (opt[k].padStart(ret[1].length, "0")))
                }
            }
            return fmt;
        },


        isChinese: function (str){
            if (escape(str).indexOf( "%u" ) < 0) return false ;
            return true ;
        }
    }
    // content end
    exportExcel.fn.init.prototype = exportExcel.fn;
    return exportExcel;
})(window);
// 初始化
var exportExcel = new exportExcel();


