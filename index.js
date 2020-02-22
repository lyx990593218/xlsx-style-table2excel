/**
 * 凌晨客车上报列表页交互控制
 * 1. 依赖provider_data_info_report.js
 */
function index() {
    this.init();
}

index.prototype = {
    /**
     * 生命周期
     */
    init: function () {
        var that = this;
        that.initUICode();              // 初始化UI的代码
    },
    /**
     * 方法
     */
    // UI code
    initUICode: function () {
        var that = this;
        that.getStasticData();
    },

    // 获取列表数据,并处理到UI
    getStasticData: function () {
        var that = this;
        // 查询列表
        $.ajax({
            type: "get",//请求方式
            url: "data.json",//地址，就是json文件的请求路径
            dataType: "json",//数据类型可以为 text xml json  script  jsonp
            async: false,
            success: function(data){
                that.stasticOfProvinceTableIns = layui.table.render({
                    elem: '#stasticTable_zxbig'
                    , data: data
                    , height: 'full'
                    , limit: data.length + 1 // 必须要设置，否则只显示10条数据
                    , cols:[
                        [
                            {rowspan: 2, align: 'center', title: '路段公司', field: 'short_name'},
                            {rowspan: 1, align: 'center', title: '收费站入口', colspan: 2},
                            {rowspan: 1, align: 'center', title: '收费站出口', colspan: 2},
                            {rowspan: 1, align: 'center', title: '入区休息', colspan: 2},
                            /*{
                                rowspan: 2, align: 'center', title: '合计车辆'
                                ,templet: function (item) {
                                    var sum = 0;
                                    if (item.col011) {
                                        sum += parseInt(item.col011);
                                    }
                                    if (item.col010) {
                                        sum += parseInt(item.col010);
                                    }
                                    if (item.col021) {
                                        sum += parseInt(item.col021);
                                    }
                                    if (item.col020) {
                                        sum += parseInt(item.col020);
                                    }
                                    if (item.col10) {
                                        sum += parseInt(item.col10);
                                    }
                                    if (item.col11) {
                                        sum += parseInt(item.col11);
                                    }
                                    if (sum === 0) {
                                        return '';
                                    }
                                    return sum;
                                }
                            },*/
                        ],
                        [
                            {align: 'center', title: '省内', field: 'type011'},
                            {align: 'center', title: '省外', field: 'type010'},
                            {align: 'center', title: '省内', field: 'type021'},
                            {align: 'center', title: '省外', field: 'type020'},
                            {align: 'center', title: '省内', field: 'type11'},
                            {align: 'center', title: '省外', field: 'type10'},
                        ]
                    ]
                    , page: false
                    , toolbar: '#toolbar'
                    , defaultToolbar: ['']
                });

                that.setListOptionBtn();
            }
        });
    },

    // 绑定事件操作按钮事件
    setListOptionBtn:function(){
        var that = this;

        layui.table.on('toolbar(stastic)', function (obj) {
            switch (obj.event){
                case 'export':
                    exportExcel.tableToExcel(
                        {
                            lay_id: 'stasticTable_zxbig',
                            configTableBefore: function () {
                                var headerColSpanNum = exportExcel.headerColSpanNum;
                                var filename = '统计表'
                                var orgname = 'org_name';

                                var reportdate = exportExcel.dateFormat('YYYY年mm月dd日 HH时', new Date());

                                var firstTrHtml = '<tr><th>1.0</th><th rowspan="2" colspan="' + (headerColSpanNum - 2) + '">' + filename + '</th><th></th></tr><tr></tr>';
                                var secendTrHtml = '<tr><th colspan="' + Math.floor(headerColSpanNum/2) + '">汇总单位：' + orgname + '</th>' +
                                    '<th colspan="' + Math.ceil(headerColSpanNum/2) + '">填报时间：' + reportdate + '</th></tr>';
                                $('#exportTable').append(firstTrHtml);
                                $('#exportTable').append(secendTrHtml);
                            },
                            filename: '统计表',
                            hpxStartRowsIndex: 3,
                            wchStartRowsIndex: 5,
                            renderSheet: function () {
                                var arr = exportExcel.arr;
                                var hpx = exportExcel.hpx;
                                var wch = exportExcel.wch;
                                var sheet = exportExcel.sheet;
                                var headerColSpanNum = exportExcel.headerColSpanNum;

                                hpx[0] = {hpx: 40};
                                hpx[1] = {hpx: 40};
                                hpx[2] = {hpx: 25};

                                //配置第一行
                                if (sheet['A' + Number(0 + 1)]) {
                                    sheet['A' + Number(0 + 1)].s = {
                                        font: {
                                            name: '宋体',
                                            sz: 12,
                                            bold: false,
                                            underline: false,
                                            color: { rgb: "FFFFFF" }
                                        },
                                        alignment: {horizontal: "center", vertical: "center", wrapText: true},
                                    };
                                }

                                if (sheet['B' + Number(0 + 1)]) {
                                    sheet['B' + Number(0 + 1)].s = {
                                        font: {
                                            name: '宋体',
                                            sz: 24,
                                            bold: true,
                                            underline: false,
                                        },
                                        alignment: {horizontal: "center", vertical: "center", wrapText: true},
                                    };
                                }

                                //配置第二行
                                if (sheet['A' + Number(2 + 1)]) {
                                    sheet['A' + Number(2 + 1)].s = {
                                        font: {
                                            name: '宋体',
                                            sz: 12,
                                            bold: false,
                                            underline: false,
                                        },
                                        alignment: {horizontal: "left", vertical: "center", wrapText: true},
                                    };
                                }

                                if (sheet[arr[Math.floor(headerColSpanNum/2)] + Number(2 + 1)]) {
                                    sheet[arr[Math.floor(headerColSpanNum/2)] + Number(2 + 1)].s = {
                                        font: {
                                            name: '宋体',
                                            sz: 12,
                                            bold: false,
                                            underline: false,
                                        },
                                        alignment: {horizontal: "right", vertical: "center", wrapText: true},
                                    };
                                }
                                //配置第二行
                                wch[0] = {wch: 30};
                                wch[5] = {wch: 30};

                            },
                            totalWch: 160
                        });
                    break;
            }

        });
    }
};

var index = new index(); // 引入即为启动页面
