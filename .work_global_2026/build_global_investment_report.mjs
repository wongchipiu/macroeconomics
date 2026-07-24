import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";
import fs from "fs";
import path from "path";

const outputDir = "D:/Users/bruce01.huang/Documents/宏观经济/outputs/global_2026_07_24";
const outputPath = path.join(outputDir, "全球经济前20国_产业投资量化分析_2026-07-24.xlsx");
const previewDir = "D:/Users/bruce01.huang/Documents/宏观经济/.work_global_2026/previews";
fs.mkdirSync(outputDir, { recursive: true });
fs.mkdirSync(previewDir, { recursive: true });

const colors = {
  navy: "#17365D",
  blue: "#4472C4",
  lightBlue: "#D9EAF7",
  paleBlue: "#EAF3F8",
  green: "#70AD47",
  darkGreen: "#548235",
  paleGreen: "#E2F0D9",
  amber: "#F4B183",
  paleAmber: "#FFF2CC",
  red: "#C00000",
  paleRed: "#FCE4D6",
  gray: "#7F8C8D",
  lightGray: "#E7E6E6",
  white: "#FFFFFF",
  black: "#222222",
  purple: "#7030A0",
};

const scoreWeights = {
  growth: 0.25,
  structure: 0.20,
  macro: 0.15,
  access: 0.15,
  valuation: 0.15,
  positioning: 0.10,
};

const countries = [
  {
    country: "中国台湾", g26: 9.64, g27: null, structure: 10.0, macro: 7.0, access: 7.0,
    pe: 31.23, yield: 0.0282, ytd: 0.7086, ret1y: 0.9891, vol: 0.4392, price: 99.38,
    etf: "EWT", macroSource: "M4", etfSource: "E1", target: 0.04,
    industries: "先进制程、AI服务器、封装测试、电子代工",
    catalyst: "2026Q1 GDP +14.55%；AI硬件出口、2nm/先进封装扩产",
    risk: "IT权重约73%、P/E约31倍；地缘与单一产业集中",
    stocks: "TSM；2317.TW 鸿海；2454.TW 联发科",
    condition: "仅回撤买：距当前价-12%/-20%分批；硬件三国合计≤12%",
  },
  {
    country: "越南", g26: 6.8, g27: 7.1, structure: 9.0, macro: 6.0, access: 4.0,
    pe: null, yield: 0.0020, ytd: -0.0141, ret1y: 0.3456, vol: 0.2071, price: 16.66,
    etf: "VNM", macroSource: "M5", etfSource: "E2", target: 0.05,
    industries: "电子制造迁移、港口物流、银行、工业园区",
    catalyst: "FDI与出口升级；AI相关出口占GDP比重快速提升；资本市场改革",
    risk: "进口依赖、汇率/流动性；VNM中IT仅约2.4%，映射AI主线偏弱",
    stocks: "FPT.VN；VCB.VN；HPG.VN",
    condition: "ETF≤组合3.25%；个股合计≤越南仓位30%，分40/30/30三段",
  },
  {
    country: "印度", g26: 6.4, g27: 6.7, structure: 9.0, macro: 6.5, access: 8.0,
    pe: 22.66, yield: 0.0000, ytd: -0.0909, ret1y: -0.1139, vol: 0.1435, price: 47.73,
    etf: "INDA", macroSource: "M1", etfSource: "E3", target: 0.09,
    industries: "金融、数字基础设施、工业资本品、医药、消费",
    catalyst: "高潜在增速、信贷深化、制造业本土化、公共资本开支",
    risk: "估值仍高、财政与油价敏感、监管及外资流波动",
    stocks: "HDB；IBN；INFY",
    condition: "P/E≤23倍可按40/30/30建仓；若油价>100美元则减1个百分点",
  },
  {
    country: "韩国", g26: 2.6, g27: 2.5, structure: 9.5, macro: 8.0, access: 9.0,
    pe: 22.83, yield: 0.0117, ytd: 1.0580, ret1y: 1.8851, vol: 0.7964, price: 169.39,
    etf: "EWY", macroSource: "M1", etfSource: "E4", target: 0.03,
    industries: "HBM存储、先进制造、船舶、核电设备、汽车",
    catalyst: "AI存储与电力设备订单；出口周期上行",
    risk: "ETF年内约+106%、1年约+189%；波动和周期回撤风险极高",
    stocks: "005930.KS 三星电子；000660.KS SK海力士；267260.KS HD现代电气",
    condition: "不追涨；先25%，回撤-15%加35%，回撤-25%加40%",
  },
  {
    country: "马来西亚", g26: 4.7, g27: 4.3, structure: 8.5, macro: 7.5, access: 6.5,
    pe: 15.06, yield: 0.0359, ytd: 0.0055, ret1y: 0.1550, vol: 0.1259, price: 27.81,
    etf: "EWM", macroSource: "M1", etfSource: "E5", target: 0.05,
    industries: "数据中心、电网、半导体封测、金融、棕榈油",
    catalyst: "已签数据中心项目约5.9GW；AI基础设施与电网投资扩张",
    risk: "电力/水资源瓶颈、项目兑现、商品与外资流波动",
    stocks: "5347.KL Tenaga；6742.KL YTL Power；0166.KL Inari",
    condition: "P/E≤16倍、股息率≥3%维持超配；数据中心组合仓位≤3.25%",
  },
  {
    country: "美国", g26: 2.3, g27: 2.2, structure: 10.0, macro: 6.0, access: 10.0,
    pe: 27.40, yield: 0.0107, ytd: 0.1107, ret1y: 0.2316, vol: 0.1516, price: 365.37,
    etf: "VTI", macroSource: "M1", etfSource: "E6", target: 0.16,
    industries: "AI算力/软件、电网、国防、医疗创新、资本市场",
    catalyst: "AI资本开支与生产率外溢；全球最深流动性和创新生态",
    risk: "VTI P/E约27.4倍；财政赤字、期限溢价与AI估值集中",
    stocks: "MSFT；AVGO；VRT",
    condition: "作为核心仓；估值>30倍时暂停增配，回撤8%/15%再加仓",
  },
  {
    country: "印度尼西亚", g26: 5.0, g27: 5.1, structure: 7.5, macro: 6.0, access: 6.0,
    pe: 9.80, yield: 0.0330, ytd: -0.3853, ret1y: -0.3404, vol: 0.3246, price: 12.40,
    etf: "EIDO", macroSource: "M1", etfSource: "E7", target: 0.04,
    industries: "镍/电池材料、银行、通信、消费、基建",
    catalyst: "约5%潜在增速、资源下游化、人口与金融渗透率",
    risk: "政策/治理、商品价格、货币与ETF流动性；深跌可能是价值陷阱",
    stocks: "BBCA.JK；TLKM.JK；ANTM.JK",
    condition: "P/E<11倍分批；单次不超过目标仓40%，止损用基本面而非价格",
  },
  {
    country: "新加坡", g26: 3.0, g27: null, structure: 8.5, macro: 9.5, access: 9.0,
    pe: 18.42, yield: 0.0377, ytd: 0.0957, ret1y: 0.1939, vol: 0.1665, price: 31.44,
    etf: "EWS", macroSource: "M3", etfSource: "E8", target: 0.04,
    industries: "金融、港口物流、数据中心、国防科技、REIT",
    catalyst: "MTI 2026增速区间2%-4%；区域资金与供应链枢纽",
    risk: "地产/银行集中、外需依赖、能源输入成本",
    stocks: "D05.SI DBS；S63.SI ST Engineering；BN4.SI Keppel",
    condition: "股息率≥3.5%、P/E≤19倍可作防守核心；仓位4%",
  },
  {
    country: "中国", g26: 4.6, g27: 4.1, structure: 9.0, macro: 6.5, access: 6.5,
    pe: 13.68, yield: 0.0205, ytd: -0.1465, ret1y: -0.0524, vol: 0.2172, price: 53.35,
    etf: "MCHI", macroSource: "M1", etfSource: "E9", target: 0.07,
    industries: "AI应用、云、电动车/电池、电网、工业自动化",
    catalyst: "增长韧性、产业链完整、估值低于多数科技市场",
    risk: "地产/需求、政策与地缘、ADR/境外上市结构风险",
    stocks: "0700.HK 腾讯；9988.HK 阿里巴巴；1211.HK 比亚迪",
    condition: "P/E<14倍可超配；若盈利预期连续2季下修则减2个百分点",
  },
  {
    country: "波兰", g26: 3.4, g27: 2.4, structure: 7.5, macro: 6.5, access: 7.0,
    pe: 15.54, yield: 0.0361, ytd: 0.1105, ret1y: 0.2501, vol: 0.2066, price: 40.78,
    etf: "EPOL", macroSource: "M1", etfSource: "E10", target: 0.04,
    industries: "国防、银行、近岸制造、软件、基建",
    catalyst: "欧盟资金、国防支出、供应链近岸化、居民需求",
    risk: "俄乌地缘、财政/通胀、市场集中与货币波动",
    stocks: "PKO.WA；DNP.WA；CDR.WA",
    condition: "P/E≤16倍且股息率>3%维持；地缘升级减半",
  },
  {
    country: "丹麦", g26: 2.5, g27: 1.5, structure: 8.5, macro: 9.5, access: 7.0,
    pe: 16.51, yield: 0.0310, ytd: -0.0182, ret1y: -0.0008, vol: 0.1535, price: 111.18,
    etf: "EDEN", macroSource: "M7", etfSource: "E11", target: 0.02,
    industries: "创新药、航运物流、风电、工业自动化",
    catalyst: "医药出口与高附加值制造；财政和制度韧性",
    risk: "单一药企/医疗权重高、产品管线与定价政策",
    stocks: "NVO；DSV.CO；VWS.CO",
    condition: "医疗仓位折算至总组合≤2%；P/E≤18倍再增配",
  },
  {
    country: "瑞典", g26: 1.9, g27: 2.5, structure: 8.5, macro: 8.5, access: 8.0,
    pe: 18.15, yield: 0.0361, ytd: 0.0389, ret1y: 0.1305, vol: 0.1828, price: 49.70,
    etf: "EWD", macroSource: "M8", etfSource: "E12", target: 0.03,
    industries: "国防、工业自动化、电信设备、金融",
    catalyst: "国防/公共投资、内需修复、资本品出口",
    risk: "地产/家庭杠杆、欧洲周期、克朗波动",
    stocks: "SAAB-B.ST；ATCO-A.ST；INVE-B.ST",
    condition: "P/E≤19倍、盈利上修时持有；地产信用利差扩大则减仓",
  },
  {
    country: "挪威", g26: 1.3, g27: 1.2, structure: 7.5, macro: 10.0, access: 7.0,
    pe: 14.92, yield: 0.0536, ytd: 0.1449, ret1y: 0.1910, vol: 0.2435, price: 34.30,
    etf: "ENOR", macroSource: "M9", etfSource: "E13", target: 0.04,
    industries: "油气、海工、国防、海产、主权财富配置",
    catalyst: "能源现金流、经常账户与财政缓冲、欧洲能源安全",
    risk: "油价反转、碳转型、ETF规模较小",
    stocks: "EQNR；KOG.OL；AKRBP.OL",
    condition: "股息率>5%维持；油价跌破70美元且盈利下修时减1.5个百分点",
  },
  {
    country: "荷兰", g26: 1.0, g27: 1.3, structure: 9.0, macro: 8.0, access: 8.0,
    pe: 20.26, yield: 0.0423, ytd: 0.2428, ret1y: 0.3510, vol: 0.2547, price: 67.10,
    etf: "EWN", macroSource: "M1", etfSource: "E14", target: 0.03,
    industries: "半导体设备、物流、工业技术、金融",
    catalyst: "AI设备资本开支、欧洲制造升级、全球贸易枢纽",
    risk: "半导体设备集中、出口管制、中国资本开支周期",
    stocks: "ASML；ASM.AS；BESI.AS",
    condition: "P/E≤20倍附近小仓；1年涨幅>35%不加仓",
  },
  {
    country: "西班牙", g26: 2.1, g27: 1.8, structure: 7.0, macro: 7.0, access: 8.0,
    pe: 15.68, yield: 0.0282, ytd: 0.1213, ret1y: 0.3856, vol: 0.1751, price: 59.02,
    etf: "EWP", macroSource: "M1", etfSource: "E15", target: 0.04,
    industries: "银行、公用事业、可再生能源、旅游、国防电子",
    catalyst: "欧元区领先增长、旅游与服务出口、能源转型投资",
    risk: "ETF金融+公用事业权重高、利率与监管风险",
    stocks: "SAN；IBE.MC；IDR.MC",
    condition: "P/E≤16倍持有；银行净息差转弱时从4%降至3%",
  },
  {
    country: "巴西", g26: 2.4, g27: 2.2, structure: 7.5, macro: 5.5, access: 7.5,
    pe: 10.70, yield: 0.0405, ytd: 0.0944, ret1y: 0.2547, vol: 0.2133, price: 36.17,
    etf: "EWZ", macroSource: "M1", etfSource: "E16", target: 0.06,
    industries: "金融、矿业、油气、农业、数字金融",
    catalyst: "低估值/高股息、商品与降息周期、金融渗透率",
    risk: "财政规则、雷亚尔、政治、商品价格",
    stocks: "NU；VALE；PBR",
    condition: "P/E≤11倍、股息率≥4%超配；财政风险溢价急升则减2个百分点",
  },
  {
    country: "瑞士", g26: 1.1, g27: 1.5, structure: 8.5, macro: 10.0, access: 9.0,
    pe: 24.30, yield: 0.0176, ytd: 0.0670, ret1y: 0.1653, vol: 0.1494, price: 61.87,
    etf: "EWL", macroSource: "M10", etfSource: "E17", target: 0.04,
    industries: "制药、精密工业、财富管理、食品消费",
    catalyst: "避险货币、低通胀、高端制造与全球医疗现金流",
    risk: "估值偏高、瑞郎过强、药品定价政策",
    stocks: "NVS；ABB；RHHBY",
    condition: "防守仓；P/E>25倍不追，风险资产回撤时作为再平衡资金源",
  },
  {
    country: "澳大利亚", g26: 1.9, g27: 1.8, structure: 7.5, macro: 8.0, access: 9.0,
    pe: 21.39, yield: 0.0298, ytd: 0.0886, ret1y: 0.1042, vol: 0.1579, price: 28.49,
    etf: "EWA", macroSource: "M11", etfSource: "E18", target: 0.04,
    industries: "矿业、银行、关键矿产、医疗、能源转型",
    catalyst: "商品现金流、养老金资本池、关键矿产与电网投资",
    risk: "中国需求、房地产/家庭债务、通胀与利率",
    stocks: "BHP；CSL.AX；MQG.AX",
    condition: "P/E≤21倍附近分批；材料+金融总暴露计入组合周期仓上限",
  },
  {
    country: "日本", g26: 0.6, g27: 0.7, structure: 8.5, macro: 6.5, access: 10.0,
    pe: 18.91, yield: 0.0379, ytd: 0.1665, ret1y: 0.2989, vol: 0.2675, price: 91.94,
    etf: "EWJ", macroSource: "M1", etfSource: "E19", target: 0.06,
    industries: "机器人、半导体设备、重工、电网、汽车",
    catalyst: "公司治理改革、名义增长、自动化与资本效率提升",
    risk: "日元、利率正常化、能源进口、出口周期",
    stocks: "TM；6501.T 日立；8035.T 东京电子",
    condition: "P/E≤19倍持有；日元单季升值>10%时降低出口股比例",
  },
  {
    country: "墨西哥", g26: 1.2, g27: 1.9, structure: 7.5, macro: 6.0, access: 8.0,
    pe: 13.60, yield: 0.0329, ytd: 0.1018, ret1y: 0.2850, vol: 0.1993, price: 75.00,
    etf: "EWW", macroSource: "M1", etfSource: "E20", target: 0.03,
    industries: "近岸制造、机场、消费、建材、金融",
    catalyst: "北美供应链重构、物流/工业园、人口与消费",
    risk: "美墨贸易规则、治安/水电、比索与财政",
    stocks: "FMX；PAC；CX",
    condition: "P/E≤14倍持有；贸易规则恶化则仓位从3%降至1.5%",
  },
];

function calcValuationScore(c) {
  if (c.pe === null || c.pe === undefined) return 5;
  return Math.min(10, Math.max(0, 8 - (c.pe - 12) / 3 + c.yield * 50));
}
function calcPositioningScore(c) {
  return Math.max(0, 10 - Math.abs(c.ytd) * 100 / 15 - Math.abs(c.vol * 100 - 18) / 10);
}
function calcComposite(c) {
  const avgGrowth = c.g27 === null ? c.g26 : c.g26 * 0.6 + c.g27 * 0.4;
  const growthScore = Math.min(10, Math.max(0, 2 + avgGrowth));
  return growthScore * scoreWeights.growth
    + c.structure * scoreWeights.structure
    + c.macro * scoreWeights.macro
    + c.access * scoreWeights.access
    + calcValuationScore(c) * scoreWeights.valuation
    + calcPositioningScore(c) * scoreWeights.positioning;
}

for (const c of countries) c.scriptScore = calcComposite(c);
countries.sort((a, b) => b.scriptScore - a.scriptScore);

const sources = [
  ["M1", "宏观", "多国", "2026/2027 GDP预测", "2026-07", "IMF", "https://www.imf.org/en/publications/weo/issues/2026/07/08/world-economic-outlook-update-july-2026?cid=ca-com-homepage-WEOET2026004", "WEO 2026年7月更新；世界增长3.0%/3.4%"],
  ["M2", "宏观", "全球", "2026全球增长与风险", "2026-06", "World Bank", "https://www.worldbank.org/en/news/press-release/2026/06/11/global-economic-prospects-june-2026-press-release", "世界银行2026年6月GEP；全球增长2.5%"],
  ["M3", "宏观", "新加坡", "2026 GDP区间2%-4%", "2026-05", "Singapore MTI", "https://www.mti.gov.sg/newsroom/mti-maintains-2026-gdp-growth-forecast-at--2-0-to-4-0-per-cent-/", "模型用区间中点3.0%；2027留空"],
  ["M4", "宏观", "中国台湾", "2026 GDP 9.64%", "2026-05", "DGBAS", "https://eng.dgbas.gov.tw/News_Content.aspx?n=4438&s=236299", "2026Q1 GDP +14.55%；2027留空"],
  ["M5", "宏观", "越南", "2026/2027 GDP", "2026-05", "World Bank", "https://www.worldbank.org/en/news/press-release/2026/05/15/viet-nam-s-economy-remains-resilient-but-sustained-reforms-are-key-to-navigating-heightened-uncertainty-wb", "2026 6.8%；统计附录2027 7.1%"],
  ["M6", "产业", "马来西亚", "数据中心/电力", "2026-04", "World Bank", "https://documents1.worldbank.org/curated/en/099040726112040954/pdf/P512544-85a8b56a-bd86-42ba-b0c7-f52661ae2879.pdf", "已签数据中心项目约5.9GW；2030潜在12.9GW"],
  ["M7", "宏观", "丹麦", "2026/2027 GDP", "2026-06", "OECD", "https://www.oecd.org/en/publications/oecd-economic-outlook-volume-2026-issue-1_2d1956f0-en/full-report/denmark_6f3418c0.html", "2.5%/1.5%"],
  ["M8", "宏观", "瑞典", "2026/2027 GDP", "2026-06", "OECD", "https://www.oecd.org/en/publications/oecd-economic-outlook-volume-2026-issue-1_2d1956f0-en/full-report/sweden_079b157b.html", "1.9%/2.5%"],
  ["M9", "宏观", "挪威", "2026/2027 GDP", "2026-06", "OECD", "https://www.oecd.org/en/about/news/press-releases/2026/06/norway-should-boost-productivity-and-limit-public-spending-to-foster-strong-and-resilient-growth.html", "总GDP 1.3%/1.2%；经常账户缓冲大"],
  ["M10", "宏观", "瑞士", "2026/2027 GDP", "2026-06", "OECD", "https://www.oecd.org/en/topics/sub-issues/economic-surveys/switzerland-economic-snapshot.html", "1.1%/1.5%"],
  ["M11", "宏观", "澳大利亚", "2026/2027 GDP", "2026-06", "OECD", "https://www.oecd.org/en/topics/sub-issues/economic-surveys/Australia-Economic-Snapshot.html", "1.9%/1.8%"],
  ["I1", "市场数据", "全部ETF", "价格、收益、波动、规模", "2026-07-24", "Interactive Brokers", "https://www.interactivebrokers.com/en/trading/market-data.php", "IBKR快照；收益截至2026-06-30附近，价格为2026-07快照"],
  ["E1", "ETF", "中国台湾", "EWT估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239686/EWT", "P/E 31.23；IT约72.93%"],
  ["E2", "ETF", "越南", "VNM行业/持仓", "2026-07", "VanEck", "https://www.vaneck.com/offshore/en/investments/vietnam-etf-vnm/", "IT约2.39%，金融/地产权重高"],
  ["E3", "ETF", "印度", "INDA估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239659/INDA", "P/E 22.66"],
  ["E4", "ETF", "韩国", "EWY估值/行业", "2026-07-16", "iShares", "https://www.ishares.com/us/products/239681/EWY", "P/E 22.83；IT约49.99%"],
  ["E5", "ETF", "马来西亚", "EWM估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239669/EWM", "P/E 15.06"],
  ["E6", "ETF", "美国", "VTI估值", "2026-05-31", "Vanguard", "https://investor.vanguard.com/investment-products/etfs/profile/vti", "P/E 27.4"],
  ["E7", "ETF", "印度尼西亚", "EIDO估值/行业", "2026-07-17", "iShares", "https://www.ishares.com/us/products/239661/ishares-msci-indonesia-etf", "P/E 9.80"],
  ["E8", "ETF", "新加坡", "EWS估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239678/ishares-msci-singapore-capped-etf", "P/E 18.42"],
  ["E9", "ETF", "中国", "MCHI估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239619/MCHI", "P/E 13.68"],
  ["E10", "ETF", "波兰", "EPOL估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239676/", "P/E 15.54"],
  ["E11", "ETF", "丹麦", "EDEN估值/行业", "2026-06-08", "iShares", "https://www.ishares.com/us/products/overview-v3-ishares-fund-data?portfolioId=239621&seoSlug=ishares-msci-denmark-capped-etf", "P/E 16.51；医疗约34.67%"],
  ["E12", "ETF", "瑞典", "EWD估值/行业", "2026-07-20", "iShares", "https://www.ishares.com/us/products/239684/ishares-msci-sweden-etf", "P/E 18.15"],
  ["E13", "ETF", "挪威", "ENOR估值/行业", "2026-07-21", "iShares", "https://www.ishares.com/us/products/239673/ishares-msci-norway-capped-etf", "P/E 14.92"],
  ["E14", "ETF", "荷兰", "EWN估值/行业", "2026-07-17", "iShares", "https://www.ishares.com/us/products/239671/ishares-msci-netherlands-etf", "P/E 20.26"],
  ["E15", "ETF", "西班牙", "EWP估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239683/ishares-msci-spain-capped-etf", "P/E 15.68；金融+公用事业权重高"],
  ["E16", "ETF", "巴西", "EWZ估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239612/ishares-msci-brazil-capped-etf", "P/E 10.70"],
  ["E17", "ETF", "瑞士", "EWL估值/行业", "2026-07-20", "iShares", "https://www.ishares.com/us/products/239685/", "P/E 24.30"],
  ["E18", "ETF", "澳大利亚", "EWA估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239607/ishares-msci-australia-etf", "P/E 21.39"],
  ["E19", "ETF", "日本", "EWJ估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239665/ishares-msci-japan-etf", "P/E 18.91"],
  ["E20", "ETF", "墨西哥", "EWW估值/行业", "2026-07-22", "iShares", "https://www.ishares.com/us/products/239670/ISHARES-MSCI-SOUTH-KOREA-ETF/", "P/E 13.60；iShares页面实际产品为EWW"],
];

const wb = Workbook.create();

function styleTitle(ws, range, text, subtitleRange, subtitle) {
  ws.mergeCells(range);
  const titleCell = ws.getRange(range.split(":")[0]);
  titleCell.values = [[text]];
  titleCell.format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white, size: 18 },
    verticalAlignment: "center",
    horizontalAlignment: "left",
  };
  ws.getRange(range).format.rowHeight = 34;
  if (subtitleRange) {
    ws.mergeCells(subtitleRange);
    const subCell = ws.getRange(subtitleRange.split(":")[0]);
    subCell.values = [[subtitle]];
    subCell.format = {
      fill: colors.paleBlue,
      font: { color: colors.navy, italic: true, size: 10 },
      verticalAlignment: "center",
      wrapText: true,
    };
    ws.getRange(subtitleRange).format.rowHeight = 30;
  }
}

function styleHeader(range) {
  range.format = {
    fill: colors.blue,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: {
      top: { style: "thin", color: colors.white },
      bottom: { style: "thin", color: colors.white },
      left: { style: "thin", color: colors.white },
      right: { style: "thin", color: colors.white },
    },
  };
  range.format.rowHeight = 28;
}

function styleBody(range) {
  range.format = {
    font: { color: colors.black, size: 9 },
    verticalAlignment: "center",
    borders: {
      bottom: { style: "hair", color: "#D9D9D9" },
    },
  };
}

function addRankConditionalFormatting(ws) {
  const scoreRange = ws.getRange("P8:P27");
  scoreRange.conditionalFormats.addColorScale({
    minColor: colors.paleRed,
    midColor: colors.paleAmber,
    maxColor: colors.paleGreen,
  });
  ws.getRange("K8:K27").conditionalFormats.addDataBar({ color: colors.blue, gradient: true });
  ws.getRange("M8:M27").conditionalFormats.addColorScale({
    minColor: colors.paleGreen,
    midColor: colors.paleAmber,
    maxColor: colors.paleRed,
  });
  ws.getRange("Q8:Q27").conditionalFormats.addCustom(
    '=Q8="超配"',
    { fill: colors.paleGreen, font: { color: colors.darkGreen, bold: true } }
  );
  ws.getRange("Q8:Q27").conditionalFormats.addCustom(
    '=OR(Q8="低配",Q8="观察/回撤买入")',
    { fill: colors.paleRed, font: { color: colors.red, bold: true } }
  );
}

// Model Inputs
const inputs = wb.worksheets.add("Model Inputs");
styleTitle(inputs, "A1:C1", "模型输入与权重", "A2:C2", "蓝色字体为可调整假设；绿色字体为跨表公式。分数反映“经济质量 × 可投资性”，不是GDP增速排行榜。");
inputs.getRange("A4:C15").values = [
  ["资产配置", "权重", "说明"],
  ["20国股票篮子", 0.65, "中等风险总组合的权益卫星仓"],
  ["1-3年美债", 0.20, "示例：SGOV/BIL；降低回撤与提供再平衡资金"],
  ["黄金", 0.10, "示例：IAU；对冲地缘、通胀尾部风险"],
  ["现金", 0.05, "等待回撤分批建仓"],
  ["增长", scoreWeights.growth, "2026/27实际GDP预测；缺失2027时只用2026"],
  ["产业结构", scoreWeights.structure, "技术扩散、制造/服务链条、资本开支（0-10）"],
  ["宏观韧性", scoreWeights.macro, "财政、外部平衡、能源与政策缓冲（0-10）"],
  ["市场可投性", scoreWeights.access, "流动性、制度、ETF映射、资本准入（0-10）"],
  ["估值", scoreWeights.valuation, "ETF P/E与股息率的规则化打分"],
  ["拥挤/动量", scoreWeights.positioning, "对极端涨幅及高波动进行惩罚"],
  ["权重合计", null, "应为100%"],
];
styleHeader(inputs.getRange("A4:C4"));
styleBody(inputs.getRange("A5:C15"));
inputs.getRange("B5:B14").format.font = { color: colors.blue };
inputs.getRange("B5:B14").format.numberFormat = "0%";
inputs.getRange("B15").formulas = [["=SUM(B9:B14)"]];
inputs.getRange("B15").format.numberFormat = "0%";
inputs.getRange("B15").format.font = { color: colors.green, bold: true };
inputs.getRange("A5:A15").format.font = { bold: true, color: colors.navy };
inputs.getRange("A4").format.columnWidth = 18;
inputs.getRange("B4").format.columnWidth = 12;
inputs.getRange("C4").format.columnWidth = 60;
inputs.getRange("A17:C23").values = [
  ["评分规则", "公式/阈值", "解释"],
  ["增长分", "MIN(10,MAX(0,2+加权GDP增长))", "高增长加分，但封顶10分"],
  ["估值分", "8-(P/E-12)/3+股息率×50", "越便宜、股息越高，分数越高；封顶10分"],
  ["拥挤分", "10-|YTD|/15%-|波动率-18%|/10", "惩罚暴涨暴跌及异常波动；封底0分"],
  ["超配", "综合分≥7.4且拥挤分≥3", "经济与投资条件兼具"],
  ["标配", "综合分≥6.8", "质量可接受，按模型仓位"],
  ["观察/回撤买", "拥挤分<3", "经济好，但价格/波动已显著透支"],
];
styleHeader(inputs.getRange("A17:C17"));
styleBody(inputs.getRange("A18:C23"));
inputs.getRange("A17").format.columnWidth = 18;
inputs.getRange("B17").format.columnWidth = 36;
inputs.getRange("C17").format.columnWidth = 60;
inputs.freezePanes.freezeRows(4);

// Country Ranking
const ranking = wb.worksheets.add("Country Ranking");
styleTitle(
  ranking,
  "A1:T1",
  "全球经济发展与可投资性：前20国量化排名",
  "A2:T2",
  "截至2026-07-24。综合权重：增长25%、产业20%、宏观韧性15%、市场可投性15%、估值15%、拥挤/动量10%。这是面向可投资市场的模型排名，不等同于“国家实力榜”。"
);
ranking.getRange("A4:H6").values = [
  ["核心结论", "经济强 ≠ 股市可追；中国台湾与韩国的AI硬件景气最高，但估值/涨幅/波动使建议降为“回撤买”。", null, null, null, null, null, null],
  ["更均衡的机会", "新加坡、马来西亚、印度、中国、丹麦/瑞典/挪威；增长、估值或防守属性更平衡。", null, null, null, null, null, null],
  ["最大映射偏差", "越南宏观与AI出口强，但VNM的IT权重很低；必须把“国家故事”与“ETF持仓”分开核验。", null, null, null, null, null, null],
];
ranking.mergeCells("B4:T4");
ranking.mergeCells("B5:T5");
ranking.mergeCells("B6:T6");
ranking.getRange("A4:A6").format = { fill: colors.lightBlue, font: { bold: true, color: colors.navy } };
ranking.getRange("B4:T6").format = { fill: "#F7FBFD", wrapText: true, font: { color: colors.black, size: 9 } };
ranking.getRange("A7:T7").values = [[
  "排名", "国家/地区", "2026 GDP%", "2027 GDP%", "加权增长%", "产业结构", "宏观韧性", "市场可投性",
  "ETF P/E", "股息率", "YTD", "1年收益", "波动率", "估值分", "拥挤分", "综合分", "建议", "ETF", "宏观来源", "ETF来源"
]];
styleHeader(ranking.getRange("A7:T7"));

const rankingRows = countries.map((c) => [
  null, c.country, c.g26 / 100, c.g27 === null ? null : c.g27 / 100, null, c.structure, c.macro, c.access,
  c.pe, c.yield, c.ytd, c.ret1y, c.vol, null, null, null, null, c.etf, c.macroSource, c.etfSource
]);
ranking.getRange("A8:T27").values = rankingRows;
for (let r = 8; r <= 27; r++) {
  ranking.getRange(`A${r}`).formulas = [[`=RANK(P${r},$P$8:$P$27,0)`]];
  ranking.getRange(`E${r}`).formulas = [[`=IF(D${r}="",C${r},C${r}*0.6+D${r}*0.4)`]];
  ranking.getRange(`N${r}`).formulas = [[`=IF(I${r}="",5,MIN(10,MAX(0,8-(I${r}-12)/3+J${r}*50)))`]];
  ranking.getRange(`O${r}`).formulas = [[`=MAX(0,10-ABS(K${r})*100/15-ABS(M${r}*100-18)/10)`]];
  ranking.getRange(`P${r}`).formulas = [[
    `=MIN(10,MAX(0,2+E${r}*100))*'Model Inputs'!$B$9+F${r}*'Model Inputs'!$B$10+G${r}*'Model Inputs'!$B$11+H${r}*'Model Inputs'!$B$12+N${r}*'Model Inputs'!$B$13+O${r}*'Model Inputs'!$B$14`
  ]];
  ranking.getRange(`Q${r}`).formulas = [[`=IF(O${r}<3,"观察/回撤买入",IF(P${r}>=7.4,"超配",IF(P${r}>=6.8,"标配","低配")))`]];
}
styleBody(ranking.getRange("A8:T27"));
ranking.getRange("C8:E27").format.numberFormat = "0.0%";
ranking.getRange("J8:M27").format.numberFormat = "0.0%";
ranking.getRange("F8:H27").format.numberFormat = "0.0";
ranking.getRange("I8:I27").format.numberFormat = "0.00x";
ranking.getRange("N8:P27").format.numberFormat = "0.00";
ranking.getRange("A8:A27").format.font = { bold: true, color: colors.navy };
ranking.getRange("B8:B27").format.font = { bold: true };
ranking.getRange("F8:H27").format.font = { color: colors.blue };
ranking.getRange("A8:A27").format.font = { color: colors.green, bold: true };
ranking.getRange("E8:E27").format.font = { color: colors.green };
ranking.getRange("N8:Q27").format.font = { color: colors.green };
ranking.getRange("P8:P27").format.font = { color: colors.green, bold: true };
ranking.getRange("R8:R27").format.font = { color: colors.purple, bold: true };
ranking.getRange("S8:T27").format.font = { color: colors.gray, size: 8 };
ranking.getRange("A8:T27").format.rowHeight = 22;
const rankWidths = [8, 14, 11, 11, 12, 11, 11, 12, 10, 10, 10, 10, 10, 10, 10, 10, 16, 9, 11, 11];
for (let i = 0; i < rankWidths.length; i++) {
  ranking.getRangeByIndexes(6, i, 21, 1).format.columnWidth = rankWidths[i];
}
ranking.freezePanes.freezeRows(7);
ranking.freezePanes.freezeColumns(2);
addRankConditionalFormatting(ranking);

// Industry Thesis
const thesis = wb.worksheets.add("Industry Thesis");
styleTitle(
  thesis,
  "A1:I1",
  "20国产业链、催化剂与标的映射",
  "A2:I2",
  "ETF优先用于国家暴露；个股只作为卫星增强。个股代码用于研究清单，不代表无条件买入建议。"
);
thesis.getRange("A4:I4").values = [[
  "模型排名", "国家/地区", "核心产业", "未来2-3年催化剂", "主要风险", "宽基ETF", "个股研究清单", "模型观点", "量化执行条件"
]];
styleHeader(thesis.getRange("A4:I4"));
const thesisRows = countries.map((c, i) => [
  i + 1, c.country, c.industries, c.catalyst, c.risk, c.etf, c.stocks,
  calcPositioningScore(c) < 3 ? "经济强，价格透支：等待回撤" : (c.scriptScore >= 7.4 ? "超配" : (c.scriptScore >= 6.8 ? "标配" : "低配/观察")),
  c.condition,
]);
thesis.getRange("A5:I24").values = thesisRows;
styleBody(thesis.getRange("A5:I24"));
thesis.getRange("A5:I24").format.wrapText = true;
thesis.getRange("A5:I24").format.rowHeight = 58;
thesis.getRange("A5:B24").format.font = { bold: true, color: colors.navy };
thesis.getRange("F5:F24").format.font = { bold: true, color: colors.purple };
thesis.getRange("H5:H24").conditionalFormats.addCustom(
  '=H5="超配"',
  { fill: colors.paleGreen, font: { color: colors.darkGreen, bold: true } }
);
thesis.getRange("H5:H24").conditionalFormats.addCustom(
  '=H5="经济强，价格透支：等待回撤"',
  { fill: colors.paleRed, font: { color: colors.red, bold: true } }
);
const thesisWidths = [9, 14, 28, 38, 36, 10, 32, 22, 42];
for (let i = 0; i < thesisWidths.length; i++) thesis.getRangeByIndexes(3, i, 21, 1).format.columnWidth = thesisWidths[i];
thesis.freezePanes.freezeRows(4);
thesis.freezePanes.freezeColumns(2);

// Portfolio
const portfolio = wb.worksheets.add("Portfolio");
styleTitle(
  portfolio,
  "A1:Q1",
  "量化组合、仓位与分批建仓",
  "A2:Q2",
  "默认：人民币投资者、美元计价、3-5年、中等风险。参考本金和汇率可调整；结果会自动联动。20国权重是65%权益篮子内部权重。"
);
portfolio.getRange("A4:D9").values = [
  ["资产类别", "标的示例", "总组合权重", "人民币金额"],
  ["20国股票篮子", "下表ETF", null, null],
  ["1-3年美债", "SGOV / BIL", null, null],
  ["黄金", "IAU", null, null],
  ["现金", "货币基金/存款", null, null],
  ["合计", "", null, null],
];
styleHeader(portfolio.getRange("A4:D4"));
styleBody(portfolio.getRange("A5:D9"));
portfolio.getRange("F4:G7").values = [
  ["可调参数", "数值"],
  ["参考本金（CNY）", 1000000],
  ["USD/CNY", 7.20],
  ["权益篮子占总组合", null],
];
styleHeader(portfolio.getRange("F4:G4"));
styleBody(portfolio.getRange("F5:G7"));
portfolio.getRange("G5:G6").format.font = { color: colors.blue, bold: true };
portfolio.getRange("G5").format.numberFormat = "#,##0";
portfolio.getRange("G6").format.numberFormat = "0.00";
portfolio.getRange("G7").formulas = [["='Model Inputs'!B5"]];
portfolio.getRange("G7").format.numberFormat = "0%";
portfolio.getRange("G7").format.font = { color: colors.green, bold: true };
for (let r = 5; r <= 8; r++) {
  portfolio.getRange(`C${r}`).formulas = [[`='Model Inputs'!B${r}`]];
  portfolio.getRange(`D${r}`).formulas = [[`=C${r}*$G$5`]];
}
portfolio.getRange("C9").formulas = [["=SUM(C5:C8)"]];
portfolio.getRange("D9").formulas = [["=SUM(D5:D8)"]];
portfolio.getRange("C5:D9").format.font = { color: colors.green };
portfolio.getRange("C5:C9").format.numberFormat = "0%";
portfolio.getRange("D5:D9").format.numberFormat = "#,##0";
portfolio.getRange("A11:Q11").values = [[
  "国家/地区", "ETF", "综合分", "波动率", "风险调整分", "模型权益权重", "目标权益权重", "总组合权重",
  "1百万元金额", "首批%", "首批金额", "二批%", "二批金额", "三批%", "三批金额", "ETF参考价USD", "首批参考份额"
]];
styleHeader(portfolio.getRange("A11:Q11"));

for (let i = 0; i < countries.length; i++) {
  const c = countries[i];
  const r = 12 + i;
  const rankRow = 8 + i;
  portfolio.getRange(`A${r}`).formulas = [[`='Country Ranking'!B${rankRow}`]];
  portfolio.getRange(`B${r}`).formulas = [[`='Country Ranking'!R${rankRow}`]];
  portfolio.getRange(`C${r}`).formulas = [[`='Country Ranking'!P${rankRow}`]];
  portfolio.getRange(`D${r}`).formulas = [[`='Country Ranking'!M${rankRow}`]];
  portfolio.getRange(`E${r}`).formulas = [[`=C${r}/MAX(D${r},10%)`]];
  portfolio.getRange(`F${r}`).formulas = [[`=E${r}/SUM($E$12:$E$31)`]];
  portfolio.getRange(`G${r}`).values = [[c.target]];
  portfolio.getRange(`H${r}`).formulas = [[`=G${r}*$G$7`]];
  portfolio.getRange(`I${r}`).formulas = [[`=H${r}*$G$5`]];
  const techEntry = c.country === "中国台湾" || c.country === "韩国";
  portfolio.getRange(`J${r}`).values = [[techEntry ? 0.25 : 0.40]];
  portfolio.getRange(`K${r}`).formulas = [[`=I${r}*J${r}`]];
  portfolio.getRange(`L${r}`).values = [[techEntry ? 0.35 : 0.30]];
  portfolio.getRange(`M${r}`).formulas = [[`=I${r}*L${r}`]];
  portfolio.getRange(`N${r}`).values = [[techEntry ? 0.40 : 0.30]];
  portfolio.getRange(`O${r}`).formulas = [[`=I${r}*N${r}`]];
  portfolio.getRange(`P${r}`).values = [[c.price]];
  portfolio.getRange(`Q${r}`).formulas = [[`=ROUNDDOWN(K${r}/$G$6/P${r},0)`]];
}
styleBody(portfolio.getRange("A12:Q31"));
portfolio.getRange("A12:F31").format.font = { color: colors.green };
portfolio.getRange("G12:G31").format.font = { color: colors.blue, bold: true };
portfolio.getRange("H12:I31").format.font = { color: colors.green };
portfolio.getRange("J12:J31").format.font = { color: colors.blue };
portfolio.getRange("K12:K31").format.font = { color: colors.green };
portfolio.getRange("L12:L31").format.font = { color: colors.blue };
portfolio.getRange("M12:M31").format.font = { color: colors.green };
portfolio.getRange("N12:N31").format.font = { color: colors.blue };
portfolio.getRange("O12:O31").format.font = { color: colors.green };
portfolio.getRange("P12:P31").format.font = { color: colors.blue };
portfolio.getRange("Q12:Q31").format.font = { color: colors.green };
portfolio.getRange("C12:C31").format.numberFormat = "0.00";
portfolio.getRange("D12:D31").format.numberFormat = "0.0%";
portfolio.getRange("E12:E31").format.numberFormat = "0.0";
portfolio.getRange("F12:H31").format.numberFormat = "0.0%";
portfolio.getRange("I12:I31").format.numberFormat = "#,##0";
portfolio.getRange("J12:J31").format.numberFormat = "0%";
portfolio.getRange("K12:K31").format.numberFormat = "#,##0";
portfolio.getRange("L12:L31").format.numberFormat = "0%";
portfolio.getRange("M12:M31").format.numberFormat = "#,##0";
portfolio.getRange("N12:N31").format.numberFormat = "0%";
portfolio.getRange("O12:O31").format.numberFormat = "#,##0";
portfolio.getRange("P12:P31").format.numberFormat = "$0.00";
portfolio.getRange("Q12:Q31").format.numberFormat = "0";
portfolio.getRange("F12:G31").conditionalFormats.addDataBar({ color: colors.blue, gradient: true });
portfolio.getRange("A33:G38").values = [
  ["执行规则", "触发条件", "动作", "仓位上限", "复核频率", "失败条件", "备注"],
  ["分批建仓", "当前/回撤8%/回撤15%", "普通国家40%/30%/30%", "按目标仓", "每周", "基本面恶化", "中国台湾/韩国用25%/35%/40%且回撤阈值更深"],
  ["再平衡", "偏离目标≥2个百分点或相对偏离≥25%", "卖高买低", "单国≤权益篮子16%", "季度", "流动性异常", "税费与资本利得规则需另行评估"],
  ["AI硬件集中", "台湾+韩国+马来西亚", "合计≤权益篮子12%", "12%", "月度", "地缘/订单下修", "避免把国家篮子变成单一AI交易"],
  ["个股卫星", "仅在ETF映射不足时", "个股合计≤国家仓位30%", "单股≤总组合2%", "月度", "治理/流动性风险", "核心仓仍以ETF为主"],
  ["汇率", "非人民币资产", "未对冲外汇≤总组合50%", "50%", "季度", "人民币趋势突变", "可用分批换汇降低择时风险"],
];
styleHeader(portfolio.getRange("A33:G33"));
styleBody(portfolio.getRange("A34:G38"));
portfolio.getRange("A33:G38").format.wrapText = true;
portfolio.getRange("A34:G38").format.rowHeight = 38;
const portWidths = [14, 10, 10, 10, 12, 13, 13, 13, 15, 9, 14, 9, 14, 9, 14, 14, 14];
for (let i = 0; i < portWidths.length; i++) portfolio.getRangeByIndexes(10, i, 22, 1).format.columnWidth = portWidths[i];
portfolio.getRange("A33:G38").format.columnWidth = 22;
portfolio.getRange("B33:B38").format.columnWidth = 28;
portfolio.getRange("C33:C38").format.columnWidth = 28;
portfolio.getRange("G33:G38").format.columnWidth = 48;
portfolio.freezePanes.freezeRows(11);
portfolio.freezePanes.freezeColumns(2);

// Scenarios
const scenarios = wb.worksheets.add("Scenarios");
styleTitle(
  scenarios,
  "A1:H1",
  "12个月情景分析与风险预算",
  "A2:H2",
  "以下回报是压力测试假设，不是收益承诺。总组合=65%二十国股票+20%短债+10%黄金+5%现金。"
);
scenarios.getRange("A4:H4").values = [[
  "情景", "概率", "20国股票", "短债", "黄金", "现金", "总组合回报", "概率贡献"
]];
styleHeader(scenarios.getRange("A4:H4"));
scenarios.getRange("A5:F7").values = [
  ["基准：AI投资延续、能源冲击缓解", 0.55, 0.10, 0.04, 0.03, 0.03],
  ["上行：生产率兑现、油价回落、盈利扩散", 0.20, 0.20, 0.01, -0.04, 0.03],
  ["下行：霍尔木兹受阻/AI估值修正", 0.25, -0.22, 0.06, 0.12, 0.03],
];
for (let r = 5; r <= 7; r++) {
  scenarios.getRange(`G${r}`).formulas = [[
    `=C${r}*'Model Inputs'!$B$5+D${r}*'Model Inputs'!$B$6+E${r}*'Model Inputs'!$B$7+F${r}*'Model Inputs'!$B$8`
  ]];
  scenarios.getRange(`H${r}`).formulas = [[`=B${r}*G${r}`]];
}
scenarios.getRange("A8:H8").values = [["概率加权期望", null, null, null, null, null, null, null]];
scenarios.getRange("B8").formulas = [["=SUM(B5:B7)"]];
scenarios.getRange("G8").formulas = [["=SUM(H5:H7)"]];
styleBody(scenarios.getRange("A5:H8"));
scenarios.getRange("B5:F7").format.font = { color: colors.blue };
scenarios.getRange("G5:H8").format.font = { color: colors.green, bold: true };
scenarios.getRange("B5:H8").format.numberFormat = "0.0%";
scenarios.getRange("A10:D16").values = [
  ["风险指标", "目标/阈值", "管理动作", "说明"],
  ["目标年化波动", "约10%-13%（总组合）", "股票仓不超过65%", "基于分散化假设，非历史回测承诺"],
  ["权益压力回撤", "-30%至-40%", "总组合压力约-12%至-20%", "AI与地缘双重冲击时可能更差"],
  ["单国上限", "权益篮子16%", "美国为上限锚；其余≤9%", "防止单一市场主导"],
  ["行业主题上限", "AI硬件12%", "台湾+韩国+马来西亚合计", "按权益篮子内部权重计"],
  ["再平衡阈值", "绝对2个百分点或相对25%", "季度执行", "降低追涨杀跌"],
  ["复核红线", "GDP/盈利/政策至少2项恶化", "暂停加仓并重估", "价格下跌本身不是卖出理由"],
];
styleHeader(scenarios.getRange("A10:D10"));
styleBody(scenarios.getRange("A11:D16"));
scenarios.getRange("A10:D16").format.wrapText = true;
scenarios.getRange("A10").format.columnWidth = 24;
scenarios.getRange("B10").format.columnWidth = 26;
scenarios.getRange("C10").format.columnWidth = 34;
scenarios.getRange("D10").format.columnWidth = 48;
scenarios.getRange("A4").format.columnWidth = 46;
for (const col of ["B", "C", "D", "E", "F", "G", "H"]) scenarios.getRange(`${col}4`).format.columnWidth = 15;
scenarios.freezePanes.freezeRows(4);

// Summary
const summary = wb.worksheets.add("Summary");
styleTitle(
  summary,
  "A1:P1",
  "2026全球经济前20国：产业与投资量化总览",
  "A2:P2",
  "基准日 2026-07-24｜默认人民币投资者、美元计价、3-5年、中等风险｜结论先看“价格是否透支”，再看GDP。"
);
summary.getRange("A4:D6").values = [
  ["全球宏观", "IMF 2026", "IMF 2027", "模型权益仓"],
  ["实际GDP增长", 0.030, 0.034, null],
  ["全球通胀", 0.047, 0.039, null],
];
styleHeader(summary.getRange("A4:D4"));
styleBody(summary.getRange("A5:D6"));
summary.getRange("D5").formulas = [["='Model Inputs'!B5"]];
summary.getRange("D5").format.font = { color: colors.green, bold: true };
summary.getRange("B5:D6").format.numberFormat = "0.0%";
summary.getRange("A8:F8").values = [["排名", "国家/地区", "综合分", "2026增长", "ETF P/E", "投资建议"]];
styleHeader(summary.getRange("A8:F8"));
for (let i = 0; i < 10; i++) {
  const r = 9 + i;
  const sourceRow = 8 + i;
  summary.getRange(`A${r}`).formulas = [[`='Country Ranking'!A${sourceRow}`]];
  summary.getRange(`B${r}`).formulas = [[`='Country Ranking'!B${sourceRow}`]];
  summary.getRange(`C${r}`).formulas = [[`='Country Ranking'!P${sourceRow}`]];
  summary.getRange(`D${r}`).formulas = [[`='Country Ranking'!C${sourceRow}`]];
  summary.getRange(`E${r}`).formulas = [[`=IF('Country Ranking'!I${sourceRow}="","",'Country Ranking'!I${sourceRow})`]];
  summary.getRange(`F${r}`).formulas = [[`='Country Ranking'!Q${sourceRow}`]];
}
styleBody(summary.getRange("A9:F18"));
summary.getRange("A9:F18").format.font = { color: colors.green };
summary.getRange("B9:B18").format.font = { color: colors.navy, bold: true };
summary.getRange("C9:C18").format.numberFormat = "0.00";
summary.getRange("D9:D18").format.numberFormat = "0.0%";
summary.getRange("E9:E18").format.numberFormat = "0.00x";
summary.getRange("A22:D22").values = [["资产类别", "总组合权重", "100万元金额", "作用"]];
styleHeader(summary.getRange("A22:D22"));
const allocationLabels = [
  ["20国股票篮子", null, null, "增长与产业升级"],
  ["1-3年美债", null, null, "防守与再平衡资金"],
  ["黄金", null, null, "地缘/通胀尾部对冲"],
  ["现金", null, null, "等待回撤"],
  ["合计", null, null, ""],
];
summary.getRange("A23:D27").values = allocationLabels;
for (let i = 0; i < 4; i++) {
  const r = 23 + i;
  summary.getRange(`B${r}`).formulas = [[`='Model Inputs'!B${5 + i}`]];
  summary.getRange(`C${r}`).formulas = [[`=B${r}*1000000`]];
}
summary.getRange("B27").formulas = [["=SUM(B23:B26)"]];
summary.getRange("C27").formulas = [["=SUM(C23:C26)"]];
styleBody(summary.getRange("A23:D27"));
summary.getRange("B23:C27").format.font = { color: colors.green };
summary.getRange("B23:B27").format.numberFormat = "0%";
summary.getRange("C23:C27").format.numberFormat = "#,##0";
summary.getRange("A30:F35").values = [
  ["结论", "量化动作", "理由", "触发器", "风险上限", "重点页"],
  ["不追台湾/韩国", "仅回撤分批", "经济和AI景气最强，但涨幅/波动已显著透支", "回撤12%-25%", "AI硬件三国≤权益12%", "Country Ranking / Portfolio"],
  ["优先均衡市场", "超配/标配马来西亚、印度、中国、新加坡", "增长、估值与可投性更平衡", "估值阈值+盈利趋势", "单国≤权益9%", "Industry Thesis"],
  ["保留防守资产", "35%短债+黄金+现金", "IMF/WB均提示能源与地缘尾部风险", "季度再平衡", "股票≤总组合65%", "Scenarios"],
  ["国家故事≠ETF", "核验ETF行业权重", "越南VNM IT权重低，无法完整代表AI出口链", "个股卫星≤国家仓30%", "单股≤总组合2%", "Industry Thesis"],
  ["按规则更新", "每季度更新GDP、P/E、盈利和政策", "模型输入均可编辑，跨表公式自动更新", "偏离2个百分点或25%", "复核失败条件", "Sources & Checks"],
];
styleHeader(summary.getRange("A30:F30"));
styleBody(summary.getRange("A31:F35"));
summary.getRange("A30:F35").format.wrapText = true;
summary.getRange("A31:F35").format.rowHeight = 42;

const rankChart = summary.charts.add("bar", summary.getRange("B8:C18"));
rankChart.titleText = "前10国综合分";
rankChart.hasLegend = false;
rankChart.setPosition("H8", "P20");
rankChart.barOptions.direction = "bar";
rankChart.barOptions.grouping = "clustered";

const allocationChart = summary.charts.add("doughnut", summary.getRange("A22:B26"));
allocationChart.titleText = "总组合资产配置";
allocationChart.hasLegend = true;
allocationChart.legend.position = "right";
allocationChart.setPosition("H22", "P37");

summary.getRange("A4").format.columnWidth = 20;
summary.getRange("B4:C4").format.columnWidth = 13;
summary.getRange("D4").format.columnWidth = 15;
summary.getRange("A8").format.columnWidth = 8;
summary.getRange("B8").format.columnWidth = 15;
summary.getRange("C8:E8").format.columnWidth = 13;
summary.getRange("F8").format.columnWidth = 18;
summary.getRange("A22").format.columnWidth = 20;
summary.getRange("B22:C22").format.columnWidth = 16;
summary.getRange("D22").format.columnWidth = 30;
summary.getRange("A30").format.columnWidth = 22;
summary.getRange("B30:F30").format.columnWidth = 34;
summary.freezePanes.freezeRows(2);

// Sources & Checks
const sourceSheet = wb.worksheets.add("Sources & Checks");
styleTitle(
  sourceSheet,
  "A1:H1",
  "来源、口径与模型检查",
  "A2:H2",
  "宏观数据优先采用2026年7月IMF、2026年6月OECD/World Bank及各经济体最新官方口径；ETF估值来自发行人，价格与收益来自IBKR快照。"
);
sourceSheet.getRange("A4:D11").values = [
  ["检查项", "公式结果", "目标", "状态"],
  ["国家数量", null, 20, null],
  ["20国权益权重", null, 1, null],
  ["总资产权重", null, 1, null],
  ["情景概率", null, 1, null],
  ["综合分范围", null, "0-10", null],
  ["来源数量", null, "≥20", null],
  ["模型版本", "2026-07-24 v1.0", "最新", "信息"],
];
styleHeader(sourceSheet.getRange("A4:D4"));
sourceSheet.getRange("B5").formulas = [["=COUNTA('Country Ranking'!B8:B27)"]];
sourceSheet.getRange("B6").formulas = [["=SUM(Portfolio!G12:G31)"]];
sourceSheet.getRange("B7").formulas = [["=SUM('Model Inputs'!B5:B8)"]];
sourceSheet.getRange("B8").formulas = [["=SUM(Scenarios!B5:B7)"]];
sourceSheet.getRange("B9").formulas = [["=TEXT(MIN('Country Ranking'!P8:P27),\"0.00\")&\" - \"&TEXT(MAX('Country Ranking'!P8:P27),\"0.00\")"]];
sourceSheet.getRange("B10").formulas = [[`=${sources.length}`]];
sourceSheet.getRange("D5").formulas = [['=IF(B5=C5,"通过","检查")']];
sourceSheet.getRange("D6").formulas = [['=IF(ABS(B6-C6)<0.0001,"通过","检查")']];
sourceSheet.getRange("D7").formulas = [['=IF(ABS(B7-C7)<0.0001,"通过","检查")']];
sourceSheet.getRange("D8").formulas = [['=IF(ABS(B8-C8)<0.0001,"通过","检查")']];
sourceSheet.getRange("D9").formulas = [['=IF(AND(MIN(\'Country Ranking\'!P8:P27)>=0,MAX(\'Country Ranking\'!P8:P27)<=10),"通过","检查")']];
sourceSheet.getRange("D10").formulas = [['=IF(B10>=20,"通过","检查")']];
styleBody(sourceSheet.getRange("A5:D11"));
sourceSheet.getRange("B5:B10").format.font = { color: colors.green, bold: true };
sourceSheet.getRange("D5:D10").format.font = { color: colors.green, bold: true };
sourceSheet.getRange("B6:B8").format.numberFormat = "0.0%";
sourceSheet.getRange("D5:D10").conditionalFormats.addCustom(
  '=D5="通过"',
  { fill: colors.paleGreen, font: { color: colors.darkGreen, bold: true } }
);
sourceSheet.getRange("D5:D10").conditionalFormats.addCustom(
  '=D5="检查"',
  { fill: colors.paleRed, font: { color: colors.red, bold: true } }
);
sourceSheet.getRange("A13:H13").values = [["ID", "类别", "国家/指标", "数据项", "数据日期", "机构", "URL", "备注"]];
styleHeader(sourceSheet.getRange("A13:H13"));
sourceSheet.getRange(`A14:H${13 + sources.length}`).values = sources;
styleBody(sourceSheet.getRange(`A14:H${13 + sources.length}`));
sourceSheet.getRange(`G14:G${13 + sources.length}`).format.font = { color: colors.blue, underline: true, size: 8 };
sourceSheet.getRange(`A14:H${13 + sources.length}`).format.wrapText = true;
sourceSheet.getRange(`A14:H${13 + sources.length}`).format.rowHeight = 36;
const sourceWidths = [9, 12, 15, 24, 14, 20, 70, 50];
for (let i = 0; i < sourceWidths.length; i++) sourceSheet.getRangeByIndexes(12, i, sources.length + 1, 1).format.columnWidth = sourceWidths[i];
sourceSheet.freezePanes.freezeRows(13);

// Workbook-level polish
for (const ws of [summary, ranking, thesis, portfolio, scenarios, inputs, sourceSheet]) {
  ws.getRange("A1").format.font = { name: "Aptos Display", bold: true, color: colors.white, size: 18 };
}

// Set active sheet and export.
wb.awareness.setActiveSheetName(summary.name);
const outputBuffer = await SpreadsheetFile.exportXlsx(wb);
fs.writeFileSync(outputPath, Buffer.from(outputBuffer.data));

// Render every sheet for visual QA.
const renderTargets = [
  [summary, "summary.png", "A1:P37"],
  [ranking, "country_ranking.png", "A1:T27"],
  [thesis, "industry_thesis.png", "A1:I24"],
  [portfolio, "portfolio.png", "A1:Q38"],
  [scenarios, "scenarios.png", "A1:H16"],
  [inputs, "model_inputs.png", "A1:C23"],
  [sourceSheet, "sources_checks.png", `A1:H${13 + sources.length}`],
];
for (const [sheet, name, range] of renderTargets) {
  const image = await wb.render({ sheetName: sheet.name, range, scale: 1 });
  fs.writeFileSync(path.join(previewDir, name), Buffer.from(await image.arrayBuffer()));
}

console.log(JSON.stringify({
  outputPath,
  previewDir,
  countries: countries.map((c, i) => ({ rank: i + 1, country: c.country, score: Number(c.scriptScore.toFixed(2)), target: c.target })),
  targetWeightSum: countries.reduce((s, c) => s + c.target, 0),
}, null, 2));
