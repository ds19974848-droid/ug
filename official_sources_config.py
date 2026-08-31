# 全国各城市工程造价信息官方来源配置
# 实测验证：2026-08-02
# strategy: api(专用适配器) / browser_login(浏览器登录) / crawl(通用爬取)
# verified 只表示网络可访问；data_status 才表示是否通过价格数据级实验。

CITY_SOURCES = {
    "北京": [
        {
            "name": "北京市住建委-造价信息",
            "url": "https://zjw.beijing.gov.cn/bjjs/gczj14/zjxx/index.shtml",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "官方列表页公开月度PDF；专用适配器实测2026-07解析1,306条含税信息价，无需登录"
        }
    ],
    "天津": [
        {
            "name": "京津冀城市地下综合管廊工程造价信息（官方PDF直抓）",
            "url": "https://zfcxjs.tj.gov.cn/ztzl_70/jjjgcjjjyth/jjjgx/glgczjxx/",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "官方公开PDF；2026-05京津冀管廊专项材料价格中天津列解析84条除税价，属专项有限来源"
        }
    ],
    "石家庄": [
        {
            "name": "石家庄市工程造价信息（官方PDF直抓）",
            "url": "https://zjj.sjz.gov.cn/columns/38884fcc-b2c6-46a2-83e5-e9a2a178d559/index.html",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；自动发现最新建安材料PDF，实测解析1,459条"
        }
    ],
    "唐山": [
        {
            "name": "唐山市住建局",
            "url": "http://zhujianju.tangshan.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "太原": [
        {
            "name": "山西省各市常用建设工程材料价格信息",
            "url": "http://zjt.shanxi.gov.cn/fwzl/bzdexx/jgxx/202607/t20260720_10180486.shtml",
            "strategy": "crawl",
            "verified": True,
            "note": "2026-05/06全省扫描PDF；太原列价格OCR 35/35正确，材料名仍需水印纠错"
        }
    ],
    "呼和浩特": [
        {
            "name": "呼和浩特市建设工程造价信息（官方PDF直抓）",
            "url": "http://zfcxjsj.huhhot.gov.cn/bsfw_91/xzzx/zjxx/",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；自动发现最新造价信息PDF，2026年第3期实测解析1,809条材料含税价格"
        }
    ],
    "沈阳": [
        {
            "name": "辽宁省住建厅",
            "url": "http://zjt.ln.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "大连": [
        {
            "name": "辽宁省建设工程信息价格查询系统（大连官方平台）",
            "url": "https://fwpt.zjt.ln.gov.cn/gczj/gczj/oldJgk/api/search.xhtml?selType=dz",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "省级平台公开查询；2026-06大连市实测解析2,273条材料价格，无需登录"
        }
    ],
    "长春": [
        {
            "name": "吉林省季度建设工程价格信息（长春XLSX直抓）",
            "url": "http://xxgk.jl.gov.cn/zcbm/fgw_98022/xxgkmlqy/",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；下载省厅季度ZIP并提取长春XLSX，实测解析478条含税价格"
        }
    ],
    "哈尔滨": [
        {
            "name": "黑龙江省住建厅",
            "url": "http://zfcxjst.hlj.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "上海": [
        {
            "name": "上海市信息价（API直连）",
            "url": "https://ciac.zjw.sh.gov.cn/JGBXMGCZJInterWeb/pc/#/HyxxHynr?bmCode=003002",
            "strategy": "api",
            "verified": True,
            "note": "专用适配器，直连官方API下载XLS"
        }
    ],
    "南京": [
        {
            "name": "南京市建委",
            "url": "http://sjw.nanjing.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "含信息价关键词(jgxx)"
        }
    ],
    "苏州": [
        {
            "name": "苏州市住房和城乡建设局",
            "url": "https://zfcjj.suzhou.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "官网可访问；站内检索未发现现行材料信息价文件或查询入口"
        }
    ],
    "无锡": [
        {
            "name": "无锡市住建局",
            "url": "http://js.wuxi.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "含信息价关键词(jgxx)"
        }
    ],
    "杭州": [
        {
            "name": "杭州建设工程材料信息价动态服务",
            "url": "https://mapi.zjzwfw.gov.cn/web/mgop/gov-open/zj/2002444903/reserved/index.html#/home",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "无需登录；专用浏览器动态适配器滚动分页并只取杭州市区，2026-07实测解析1,442条含税价格"
        }
    ],
    "宁波": [
        {
            "name": "宁波市建设工程造价管理协会-建材商情版",
            "url": "https://www.nbzj.net/Book/ElectronicJournalList.aspx?CategoryId=194",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；官方列表页和PDF下载均公开，2026-07建材商情版实测解析约1.2万条含税信息价，无需登录"
        },
        {
            "name": "宁波市住建局",
            "url": "http://zjw.ningbo.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "data_status": "no_public_source",
            "note": "官网可访问；材料信息价以协会期刊为主，住建局页面不作为抓取来源"
        }
    ],
    "温州": [
        {
            "name": "温州市住建局",
            "url": "http://zjj.wenzhou.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "福州": [
        {
            "name": "福建省定额材料综合价格（省级兜底）",
            "url": "https://zjt.fujian.gov.cn/hygl/gczj/gczj/202601/t20260122_7083915.htm",
            "strategy": "crawl",
            "verified": True,
            "note": "公开XLSX含5,109条；属于全省2025定额材料综合价，不等同福州市月度信息价"
        }
    ],
    "厦门": [
        {
            "name": "厦门市住房和建设局",
            "url": "https://szjj.xm.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "正确官网可访问；官方站内搜索和福建政务服务入口未发现现行材料信息价公开数据"
        }
    ],
    "济南": [
        {
            "name": "山东省造价信息网",
            "url": "http://www.sdzjxx.com/",
            "strategy": "crawl",
            "verified": True,
            "note": "含价格关键词(price)"
        },
        {
            "name": "济南市住建局",
            "url": "http://jncc.jinan.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "青岛": [
        {
            "name": "青岛市建设工程材料价格（官方PDF直抓）",
            "url": "http://sjw.qingdao.gov.cn/cxjsj13/cxjs_95/cxjsj_gczjxx13/",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；自动发现最新PDF，2026-06实测解析236条"
        }
    ],
    "烟台": [
        {
            "name": "烟台市住建局",
            "url": "http://zjj.yantai.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "郑州": [
        {
            "name": "郑州市建设工程主要材料价格（官方PDF直抓）",
            "url": "https://zzjsj.zhengzhou.gov.cn/zjxx/index.jhtml",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；自动发现最新月度PDF，2026-06实测解析939条含税价格"
        }
    ],
    "洛阳": [
        {
            "name": "洛阳造价信息网",
            "url": "http://www.lyszj.com/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问；域名可能过期"
        }
    ],
    "武汉": [
        {
            "name": "武汉市工程造价信息发布公共服务",
            "url": "http://zwfw.hubei.gov.cn/s/web/bszn/bsznpage.html?transactCode=11420100MB0889006F342201722700001",
            "strategy": "crawl",
            "verified": True,
            "note": "湖北政务服务指南确认由武汉市自然资源和城乡建设局发布材料市场价格；实际数据入口待定位"
        }
    ],
    "长沙": [
        {
            "name": "湖南省建设工程材料价格行情资讯（长沙官方平台）",
            "url": "http://zjt.hunan.gov.cn/zjt/hnweb/xzzx/zlxx/",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "省级平台公开PDF；2026年第二期（3-4月）长沙主要材料价格实测解析104条，无需登录"
        }
    ],
    "广州": [
        {
            "name": "广州市建设工程人工、材料、施工机具价格信息（官方PDF直抓）",
            "url": "https://zfcj.gz.gov.cn/zwgk/zsdwwj/",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；自动发现最新官方通知和PDF，实测2026-06解析2,508条税前综合价格，无需登录或会员"
        }
    ],
    "深圳": [
        {
            "name": "深圳市造价信息查询系统（API直连）",
            "url": "https://zjj.sz.gov.cn/szzjxx/web/pc/index",
            "strategy": "api",
            "verified": True,
            "note": "专用适配器，分页获取含税价格"
        }
    ],
    "东莞": [
        {
            "name": "东莞市住建局",
            "url": "https://zjj.dg.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "佛山": [
        {
            "name": "佛山市绿色建材市场价格（官方PDF直抓）",
            "url": "http://fszj.foshan.gov.cn/ywxt/jsgczjfwzx/zwzt_1110045/jjyjgl/jgxx/scjg/index.html",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；仅覆盖季度绿色建材税前区间价，区间取中点并保留原值"
        }
    ],
    "珠海": [
        {
            "name": "珠海市建设工程造价信息（官方月度PDF）",
            "url": "https://zjj.zhuhai.gov.cn/zjj/hygl/zzxx/index.html",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "住建局官方月度通知PDF；2026-06实测解析805条，无需登录"
        }
    ],
    "南宁": [
        {
            "name": "南宁市住建局",
            "url": "http://zjj.nanning.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "海口": [
        {
            "name": "海南建设标准定额信息平台",
            "url": "http://www.hnjsbd.com/",
            "strategy": "crawl",
            "verified": False,
            "note": "省厅官方通知指向该平台，但静态请求和真实浏览器均持续超时"
        }
    ],
    "重庆": [
        {
            "name": "重庆市建设工程造价信息网（官方期刊PDF直抓）",
            "url": "http://www.cqsgczjxx.org/Pages/CQZJW/index.aspx",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "官方列表页需浏览器渲染但PDF公开下载；专用适配器实测2026-07解析1,534条不含税材料价格，无需登录"
        }
    ],
    "成都": [
        {
            "name": "四川造价信息网（API/需授权）",
            "url": "http://202.61.90.35:8037/jgxx.htm?code=5101",
            "strategy": "api",
            "verified": True,
            "note": "专用适配器；目录公开761条，单价需会员/登录"
        }
    ],
    "绵阳": [
        {
            "name": "绵阳市区材料价格信息（官方XLS直抓）",
            "url": "https://zjw.my.gov.cn/myszjj/c101133/list.shtml",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；自动发现最新XLS，2026-06实测解析73条不含税价格"
        }
    ],
    "贵阳": [
        {
            "name": "贵州省住建厅",
            "url": "http://zfcxjst.guizhou.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "含造价关键词(xxj)"
        }
    ],
    "昆明": [
        {
            "name": "云南主材综合价格信息（官方网页直抓）",
            "url": "https://www.ynbzde.com/catlist.html?catid=32",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；2026-06实测解析14条昆明市主材综合除税价，无需登录"
        }
    ],
    "拉萨": [
        {
            "name": "西藏住建厅",
            "url": "http://zjt.xizang.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "西安": [
        {
            "name": "西安市住建局",
            "url": "http://zjj.xa.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "兰州": [
        {
            "name": "兰州市住建局",
            "url": "http://zjj.lanzhou.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "银川": [
        {
            "name": "宁夏住建厅",
            "url": "http://jst.nx.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        },
        {
            "name": "银川市住建局",
            "url": "http://zjj.yinchuan.gov.cn/",
            "strategy": "crawl",
            "verified": True,
            "note": "可访问"
        }
    ],
    "乌鲁木齐": [
        {
            "name": "乌鲁木齐市建设工程综合价格信息（官方）",
            "url": "https://www.xjzj.com/",
            "strategy": "api",
            "verified": True,
            "data_status": "working",
            "note": "专用适配器；实测下载官方XLSX并解析563条"
        }
    ]
}


# 造价HOME当前已验证到的省市目录。城市名称使用页面展示名，适配器会在请求时
# 从对应省份页面重新读取原始 data-city 参数，避免同名城市或尾随空格导致串城。
_ZAOJIAHOME_CITY_PROVINCES = {
    "上海": "上海",
    "苏州": "江苏", "常州": "江苏", "淮安": "江苏", "连云港": "江苏",
    "南京": "江苏", "南通": "江苏", "无锡": "江苏", "徐州": "江苏",
    "盐城": "江苏", "扬州": "江苏",
    "嘉兴": "浙江", "宁波": "浙江", "衢州": "浙江", "绍兴": "浙江", "温州": "浙江",
    "安庆": "安徽", "蚌埠": "安徽", "池州": "安徽", "滁州": "安徽", "阜阳": "安徽",
    "合肥": "安徽", "淮南": "安徽", "黄山": "安徽", "马鞍山": "安徽", "铜陵": "安徽",
    "福州": "福建", "宁德": "福建", "泉州": "福建",
    "滨州": "山东", "菏泽": "山东", "泰安": "山东", "威海": "山东", "潍坊": "山东",
    "枣庄": "山东", "淄博": "山东",
    "潮州": "广东", "东莞": "广东", "佛山": "广东", "广州": "广东", "惠州": "广东",
    "茂名": "广东", "梅州": "广东", "汕头": "广东", "韶关": "广东", "阳江": "广东",
    "肇庆": "广东", "中山": "广东",
    "成都": "四川", "达州": "四川", "乐山": "四川", "泸州": "四川", "绵阳": "四川",
    "遂宁": "四川",
}

# 公开市场参考入口支持按省份和城市动态定位期刊。这里维护的是城市到
# 省份的路由信息，不代表这些城市都已有官方信息价 API；官方来源仍由
# CITY_SOURCES 中的专用适配器单独决定。
_ZAOJIAHOME_CITY_PROVINCES.update({
    **{"北京": "北京", "天津": "天津", "上海": "上海", "重庆": "重庆"},
    **{"石家庄": "河北", "太原": "山西", "呼和浩特": "内蒙古", "沈阳": "辽宁", "大连": "辽宁", "长春": "吉林", "哈尔滨": "黑龙江"},
    **{"南京": "江苏", "杭州": "浙江", "合肥": "安徽", "福州": "福建", "厦门": "福建", "济南": "山东", "青岛": "山东", "烟台": "山东"},
    **{"郑州": "河南", "洛阳": "河南", "武汉": "湖北", "长沙": "湖南", "广州": "广东", "深圳": "广东", "珠海": "广东"},
    **{"南宁": "广西", "海口": "海南", "成都": "四川", "贵阳": "贵州", "昆明": "云南", "拉萨": "西藏", "西安": "陕西", "兰州": "甘肃", "银川": "宁夏", "乌鲁木齐": "新疆"},
    **{city: "河北" for city in ("秦皇岛", "邯郸", "邢台", "保定", "张家口", "承德", "沧州", "廊坊", "衡水")},
    **{city: "山西" for city in ("大同", "阳泉", "长治", "晋城", "朔州", "晋中", "运城", "忻州", "临汾", "吕梁")},
    **{city: "内蒙古" for city in ("包头", "乌海", "赤峰", "通辽", "鄂尔多斯", "呼伦贝尔", "巴彦淖尔", "乌兰察布")},
    **{city: "辽宁" for city in ("鞍山", "抚顺", "本溪", "丹东", "锦州", "营口", "阜新", "辽阳", "盘锦", "铁岭", "朝阳", "葫芦岛")},
    **{city: "吉林" for city in ("吉林", "四平", "辽源", "通化", "白山", "松原", "白城")},
    **{city: "黑龙江" for city in ("齐齐哈尔", "鸡西", "鹤岗", "双鸭山", "大庆", "伊春", "佳木斯", "七台河", "牡丹江", "黑河", "绥化")},
    **{city: "江苏" for city in ("镇江", "泰州", "宿迁")},
    **{city: "浙江" for city in ("湖州", "金华", "台州", "丽水")},
    **{city: "安徽" for city in ("芜湖", "宿州", "六安", "亳州", "宣城")},
    **{city: "福建" for city in ("莆田", "漳州", "三明", "南平", "龙岩")},
    **{city: "江西" for city in ("南昌", "景德镇", "萍乡", "九江", "新余", "鹰潭", "赣州", "吉安", "宜春", "抚州", "上饶")},
    **{city: "山东" for city in ("临沂", "济宁", "德州", "聊城")},
    **{city: "河南" for city in ("开封", "安阳", "鹤壁", "新乡", "焦作", "濮阳", "许昌", "漯河", "三门峡", "南阳", "商丘", "信阳", "周口", "驻马店", "济源")},
    **{city: "湖北" for city in ("黄石", "十堰", "宜昌", "襄阳", "鄂州", "荆门", "孝感", "荆州", "黄冈", "咸宁", "随州")},
    **{city: "湖南" for city in ("株洲", "湘潭", "衡阳", "邵阳", "岳阳", "常德", "张家界", "益阳", "郴州", "永州", "怀化", "娄底")},
    **{city: "广东" for city in ("江门", "湛江", "清远", "揭阳", "云浮", "河源", "汕尾")},
    **{city: "广西" for city in ("柳州", "桂林", "梧州", "北海", "防城港", "钦州", "贵港", "玉林", "百色", "贺州", "河池", "来宾", "崇左")},
    **{city: "海南" for city in ("三亚", "三沙", "儋州")},
    **{city: "四川" for city in ("自贡", "攀枝花", "德阳", "广元", "内江", "南充", "眉山", "宜宾", "广安", "雅安", "巴中", "资阳")},
    **{city: "贵州" for city in ("六盘水", "遵义", "安顺", "毕节", "铜仁")},
    **{city: "云南" for city in ("曲靖", "玉溪", "保山", "昭通", "丽江", "普洱", "临沧")},
    **{city: "陕西" for city in ("铜川", "宝鸡", "咸阳", "渭南", "延安", "汉中", "榆林", "安康", "商洛")},
    **{city: "甘肃" for city in ("嘉峪关", "金昌", "白银", "天水", "武威", "张掖", "平凉", "酒泉", "庆阳", "定西", "陇南")},
    **{city: "青海" for city in ("西宁", "海东")},
    **{city: "宁夏" for city in ("石嘴山", "吴忠", "固原", "中卫")},
    **{city: "新疆" for city in ("克拉玛依", "吐鲁番", "哈密", "昌吉", "博尔塔拉", "巴音郭楞", "阿克苏", "喀什", "和田", "伊犁")},
})

# Only these cities have a verified market-reference source record in the
# catalog. Other routed cities use the same adapter dynamically and remain
# clearly marked as unverified until a real price file is parsed.
_ZAOJIAHOME_VERIFIED_MARKET_CITIES = frozenset({
    "上海", "苏州", "常州", "淮安", "连云港", "南京", "南通", "无锡", "徐州", "盐城", "扬州",
    "嘉兴", "宁波", "衢州", "绍兴", "温州", "安庆", "蚌埠", "池州", "滁州", "阜阳", "合肥",
    "淮南", "黄山", "马鞍山", "铜陵", "福州", "宁德", "泉州", "滨州", "菏泽", "泰安", "威海",
    "潍坊", "枣庄", "淄博", "潮州", "东莞", "佛山", "广州", "惠州", "茂名", "梅州", "汕头",
    "韶关", "阳江", "肇庆", "中山", "成都", "达州", "乐山", "泸州", "绵阳", "遂宁",
})


def get_zaojiahome_province(city: str) -> str:
    """Return the verified 造价HOME province for a display city name."""
    value = (city or "").strip()
    if value.endswith("市"):
        value = value[:-1]
    return _ZAOJIAHOME_CITY_PROVINCES.get(value, "")


# 江苏市级期刊已定位到造价库分发目录。该站不是政府域名，明细下载需要登录/购买，
# 因此作为参考来源保留在来源配置中，但不能进入官方自动抓取任务。
_JIANGSU_COSTKU_SOURCE_PATHS = {
    "南京": "nanjing",
    "无锡": "wuxi",
    "徐州": "xuzhou",
    "常州": "changzhou",
    "苏州": "suzhou",
    "南通": "nantong",
    "连云港": "lianyungang",
    "淮安": "huaian",
    "盐城": "yancheng",
    "扬州": "yangzhou",
    "镇江": "zhenjiang",
    "泰州": "taizhou",
    "宿迁": "suqian",
}

for _city, _path in _JIANGSU_COSTKU_SOURCE_PATHS.items():
    _market_source = {
        "name": f"{_city}市官方信息价期刊（造价库分发）",
        "url": f"https://www.costku.com/{_path}/",
        "strategy": "browser_login",
        "verified": True,
        "is_official": False,
        "data_status": "membership_required",
        "note": "已定位官方信息价期刊目录；造价库为第三方分发页面，下载明细需手机号登录并购买",
    }
    # 追加第三方登录来源，不能覆盖同一城市已有的官方来源。
    CITY_SOURCES.setdefault(_city, []).append(_market_source)

_JIANGSU_PUBLIC_SOURCES = {
    "徐州": {
        "name": "徐州市主要建筑材料市场参考价格（省总站公告）",
        "url": "http://49.77.204.6:10081/continue/xydt/001004/001004003/secondPageThird.html",
        "note": "官方公告公开月度XLS附件；2026-06实测解析1,835条含税材料价格，价格性质为市场参考价",
    },
    "连云港": {
        "name": "连云港市建筑工程材料信息价（省总站公告）",
        "url": "http://49.77.204.6:10081/continue/xydt/001004/001004007/secondPageThird.html",
        "note": "官方公告公开月度PDF附件；2026-06实测解析17页材料信息价，无需登录",
    },
}
for _city, _public in _JIANGSU_PUBLIC_SOURCES.items():
    CITY_SOURCES[_city].insert(0, {
        **_public,
        "strategy": "api",
        "verified": True,
        "is_official": True,
        "data_status": "working",
    })

# 造价HOME是公开的第三方市场参考入口。它不冒充政府官方来源，
# 不保存具体附件地址，由专用适配器按当前城市和期数自动定位公开Excel。
_ZAOJIAHOME_MARKET_SOURCE = {
    "name": "公开市场参考价",
    "url": "https://wlg.zaojiahome.com/Home",
    "strategy": "api",
    "adapter_mode": "api_compatible",
    "transport": "catalog_period_document",
    "verified": True,
    "is_official": False,
    "data_status": "working",
    "source_class": "market_reference",
    "note": "公开第三方市场参考入口；按所选城市和期数自动定位公开数据，优先采用除税单价；不等同政府官方信息价",
}


def get_public_market_source_config(city: str, province: str = "") -> dict | None:
    """Build the generic public-market source for any routed major city."""
    value = (city or "").strip().removesuffix("市")
    province = (province or "").strip() or _ZAOJIAHOME_CITY_PROVINCES.get(value, "")
    if not value or not province:
        return None
    return {
        **_ZAOJIAHOME_MARKET_SOURCE,
        "name": f"{province}{value}公开市场参考价",
        "province": province,
    }


for _city, _province in _ZAOJIAHOME_CITY_PROVINCES.items():
    if _city in _ZAOJIAHOME_VERIFIED_MARKET_CITIES:
        CITY_SOURCES.setdefault(_city, []).append(get_public_market_source_config(_city, _province))

# 数据级实验状态。只有 working 会进入“一键抓取”；其余状态用于明确说明阻塞点。
CITY_AUDIT_STATUS = {
    "北京": {"status": "working", "note": "北京市住建委官方列表页和月度PDF已接入专用适配器；2026-07实测解析1,306条含税信息价，无需登录。"},
    "天津": {"status": "working", "note": "天津住建委官网公开京津冀管廊专项造价信息PDF；专用适配器实测2026-05解析天津列84条材料除税价。该来源是管廊专项，不是天津完整月度信息价。"},
    "石家庄": {"status": "working", "note": "专用附件适配器实测成功：2026-06官方PDF，解析1,459条，无需登录。"},
    "唐山": {"status": "waf_blocked", "note": "唐山市住建局真实浏览器请求返回NWAF 405；当前阻塞是网站防火墙，不是登录或会员。"},
    "太原": {"status": "public_scanned_document", "note": "省厅公开2026-05/06全省材料价格PDF，无需登录但只有扫描图。高分辨率OCR对太原价格列35/35正确，少量材料名受水印污染，待纠错后上线。"},
    "呼和浩特": {"status": "working", "note": "专用附件适配器实测成功：2026年第3期（5-6月）官方PDF，解析1,809条材料含税价格（1,807个唯一编码，2个编码跨分类合法复用），无需登录。"},
    "沈阳": {"status": "login_required", "note": "已定位辽宁建设工程信息价格查询系统；动态页面出现登录要求，需浏览器授权实验。"},
    "大连": {"status": "working", "note": "辽宁省建设工程信息价格查询系统公开可查；专用适配器实测2026-06大连市解析2,273条材料价格，无需登录。"},
    "长春": {"status": "working", "note": "专用附件适配器实测成功：吉林省2026年第2季度ZIP内长春XLSX，解析478条含税价格，无需登录。"},
    "哈尔滨": {"status": "wrong_content", "note": "目前发现的是工程造价备案办事指南，不是材料信息价。"},
    "上海": {"status": "working", "note": "专用API实测成功：2026-07官方XLS，解析4,697条，无需登录。"},
    "南京": {"status": "pending_adapter", "note": "建委官网可访问，当前入口不是明确的信息价数据页。"},
    "苏州": {"status": "no_public_source", "note": "已改为苏州市住建局正确官网；官方站内检索未发现现行材料信息价文件或查询入口。"},
    "无锡": {"status": "pending_parser", "note": "发现市政原材料公示表，页面可读但现有字段结构解析为0条，待定制解析。"},
    "杭州": {"status": "working", "note": "浙江政务MGOP动态服务无需登录；专用浏览器适配器滚动分页并只取杭州市区，2026-07实测解析1,442条含税价格。接口总数1,460条中另有18条属于其他区县。"},
    "宁波": {"status": "working", "note": "宁波市建设工程造价管理协会官方期刊公开下载PDF；专用适配器实测2026-07建材商情版约1.2万条含税信息价，无需登录。"},
    "温州": {"status": "pending_adapter", "note": "住建局官网可访问，未定位到价格数据入口。"},
    "福州": {"status": "province_fallback", "note": "福州市住建局站内搜索实测信息价仅2008年政策文件，2026年无月度材料信息价发布；仅保留福建省厅2025定额材料综合价XLSX作为省级兜底。"},
    "厦门": {"status": "no_public_source", "note": "厦门市住房和建设局智能检索实测信息价仅人工费指数，材料价格相关结果未含当前材料名称、规格、单位、价格数据；造价站未公开现行材料信息价。"},
    "济南": {"status": "login_required", "note": "山东信息价期刊页可访问但解析为0条，页面含登录/会员入口，需授权后抓接口。"},
    "青岛": {"status": "working", "note": "专用附件适配器实测成功：2026-06官方PDF，解析236条，无需登录。"},
    "烟台": {"status": "no_current_source", "note": "官网仅定位到2020年公开价格DOCX，未找到当前月度材料价格来源。"},
    "郑州": {"status": "working", "note": "专用附件适配器实测成功：2026-06官方PDF，解析939条含税价格，无需登录。"},
    "洛阳": {"status": "wrong_url", "note": "配置的旧造价域名已变为无关网站，必须更换官方现行入口。"},
    "武汉": {"status": "pending_adapter", "note": "武汉市政府全站检索实测无信息价/材料价格结果；whjs.wuhan.gov.cn 和旧造价域名均无法连接，湖北政务办理页为动态页面，实际数据入口仍待定位。"},
    "长沙": {"status": "working", "note": "湖南省建设工程材料价格行情资讯公开PDF含长沙3-4月主要材料价格；专用适配器实测解析104条，属主要材料行情，不是长沙完整月度材料目录。"},
    "广州": {"status": "working", "note": "广州市住建局官方通知和PDF无需登录、会员或验证码；专用适配器自动发现最新附件，2026-06实测解析2,508条税前综合价格（2,506个唯一编码，2个编码合法重复使用）。"},
    "深圳": {"status": "working", "note": "专用API实测成功：2026-07分页JSON，解析1,155条，无需登录。"},
    "东莞": {"status": "pending_adapter", "note": "住建局动态官网可访问，未定位到价格接口。"},
    "佛山": {"status": "working", "note": "专用附件适配器实测成功：2026年第2季度官方绿色建材PDF可公开抓取；仅覆盖绿色建材税前区间价，不代表完整月度信息价。"},
    "珠海": {"status": "working", "note": "住建局官方月度PDF已通过生产级适配验证；2026-06实测解析805条并直接入库。"},
    "南宁": {"status": "wrong_content", "note": "发现的是商品房监管造价指标，不是材料信息价，页面解析为0条。"},
    "海口": {"status": "connection_blocked", "note": "海南省厅官方通知指向hnjsbd.com；静态请求和真实浏览器均持续超时，属于官方平台连接失效，不是登录问题。"},
    "重庆": {"status": "working", "note": "重庆市建设工程造价信息网官方期刊PDF公开下载；专用浏览器适配器发现列表并解析2026-07共1,534条不含税材料价格，无需登录。"},
    "成都": {"status": "membership_required", "note": "官方接口实测目录761条、目标期93条，单价统一返回“会员查看”。"},
    "绵阳": {"status": "working", "note": "专用附件适配器实测成功：2026-06官方XLS，解析73条不含税价格，无需登录。"},
    "贵阳": {"status": "login_required", "note": "造价专栏和国家造价平台可访问，价格平台要求登录认证。"},
    "昆明": {"status": "working", "note": "专用网页适配器实测成功：云南省住建厅指导的主材综合价格页，2026-06解析14条昆明市除税价，无需登录。该页面是有限品类综合价，不是完整材料目录。"},
    "拉萨": {"status": "connection_blocked", "note": "已定位造价数据监测平台，但实测TLS证书链失败，需浏览器继续验证。"},
    "西安": {"status": "login_required", "note": "工程造价信息页可访问但为动态登录页面，静态解析0条。"},
    "兰州": {"status": "public_document", "note": "官网公开2026年第1期造价信息PDF；107页可下载，现有解析器未识别价格行。"},
    "银川": {"status": "waf_blocked", "note": "宁夏住建厅和银川住建局真实浏览器均返回NWAF 405；当前阻塞是网站防火墙，尚不能判断数据是否公开。"},
    "乌鲁木齐": {"status": "working", "note": "专用适配器实测成功：2026-05官方XLSX，解析563条，无需登录。"},
}

for _city in _JIANGSU_COSTKU_SOURCE_PATHS:
    CITY_AUDIT_STATUS[_city] = {
        "status": "membership_required",
        "note": "已确认造价库存在该城市官方信息价期刊目录；点击下载会要求手机号验证码登录，明细还需购买，当前不能直接抓取。",
    }
CITY_AUDIT_STATUS["徐州"] = {
    "status": "working",
    "note": "省总站官方公告已验证：2026-06月度XLS附件可直接下载并解析；同时保留造价库登录/购买入口作为备用。",
}
CITY_AUDIT_STATUS["连云港"] = {
    "status": "working",
    "note": "省总站官方公告已验证：2026-06月度PDF附件可直接下载并解析；同时保留造价库登录/购买入口作为备用。",
}
# 公开市场入口的二次验证结果（2026-08-28）。这里记录的是数据级结果，
# 不把“城市页能打开”误报为“已经解析出价格”。
CITY_AUDIT_STATUS["苏州"] = {
    "status": "working",
    "note": "公开市场参考入口已定位苏州城市页和2026-07期；公开XLS可下载，实测解析93条名称、单位和价格数据。来源为市场参考价，不等同政府官方信息价。",
}
CITY_AUDIT_STATUS["东莞"] = {
    "status": "working",
    "note": "公开市场参考入口已定位东莞城市页和2026-05期；公开PDF可下载，实测解析65条名称、单位和价格数据。来源为市场参考价，不等同政府官方信息价。",
}
CITY_AUDIT_STATUS["南京"] = {
    "status": "working",
    "note": "公开市场参考入口已验证：2026-07公开PDF可下载并解析1,465条人材机价格，无需登录；来源为市场参考价，不等同政府官方信息价。",
}
CITY_AUDIT_STATUS["无锡"] = {
    "status": "public_document",
    "note": "公开市场参考入口已定位无锡城市页和2026-06期，公开PDF可下载；当前PDF解析耗时较长，暂不宣称已完成价格解析，后续需优化解析器。",
}
CITY_AUDIT_STATUS["上海"] = {
    "status": "working",
    "note": "官方API已可用；公开市场参考入口同步验证2026-04 XLS，解析3,127条人材机价格，无需登录。",
}
CITY_AUDIT_STATUS["徐州"] = {
    "status": "working",
    "note": "官方公告已可用；公开市场参考入口同步验证2026-05 XLS，解析1,852条人材机价格，无需登录。",
}
CITY_AUDIT_STATUS["茂名"] = {
    "status": "working",
    "note": "公开市场参考入口已验证：2026-06 XLS解析39条人材机价格，无需登录；来源为市场参考价，不等同政府官方信息价。",
}
CITY_AUDIT_STATUS["淮安"] = {
    "status": "working",
    "note": "公开市场参考入口已验证：2026-07公开PDF解析718条人材机价格，无需登录；来源为市场参考价，不等同政府官方信息价。",
}
CITY_AUDIT_STATUS["合肥"] = {
    "status": "working",
    "note": "公开市场参考入口已验证：2026-08公开PDF解析1,309条人材机价格，无需登录；来源为市场参考价，不等同政府官方信息价。",
}
CITY_AUDIT_STATUS["肇庆"] = {
    "status": "working",
    "note": "公开市场参考入口已验证：2026-07公开PDF解析569条人材机价格，无需登录；来源为市场参考价，不等同政府官方信息价。",
}

for _city, _sources in CITY_SOURCES.items():
    _city_status = CITY_AUDIT_STATUS.get(_city, {}).get("status", "pending_adapter")
    for _source in _sources:
        _source.setdefault("data_status", _city_status)


def get_city_source_config(city: str, url: str = ""):
    """Return the canonical source configuration for a city and URL."""
    city = (city or "").strip()
    normalized_url = (url or "").strip().rstrip("/")
    for source in CITY_SOURCES.get(city, []):
        if not normalized_url or source.get("url", "").strip().rstrip("/") == normalized_url:
            return source
    return None


def is_api_compatible_source(config: dict | None) -> bool:
    """Return whether a configured source uses the unified API-style adapter."""
    return bool(
        config
        and config.get("strategy") == "api"
        and config.get("adapter_mode") == "api_compatible"
    )


def get_city_audit_status(city: str) -> dict:
    return CITY_AUDIT_STATUS.get(city, {
        "status": "pending_adapter",
        "note": "尚未完成价格数据级实验。",
    })
