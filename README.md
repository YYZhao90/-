<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>250kW/1MWh储能峰谷套利IRR测算工具</title>
    <style>
        * {box-sizing: border-box;font-family: "Microsoft Yahei",sans-serif;}
        body {max-width:1200px;margin:20px auto;padding:0 20px;background:#f5f7fa;}
        .card {background:#fff;padding:20px;border-radius:10px;box-shadow:0 2px 12px #e2e6eb;margin-bottom:20px;}
        h2 {color:#2d3748;border-left:4px solid #3182ce;padding-left:10px;}
        .grid-row {display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin:10px 0;}
        .grid-two {display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:10px 0;}
        label {display:block;margin-bottom:4px;color:#4a5568;font-size:14px;}
        input {width:100%;padding:8px 10px;border:1px solid #cbd5e0;border-radius:6px;font-size:15px;}
        button {padding:10px 24px;border:none;border-radius:6px;cursor:pointer;font-size:15px;margin-right:10px;}
        #calcBtn {background:#3182ce;color:#fff;}
        #resetBtn {background:#a0aec0;color:#fff;}
        #exportBtn {background:#38a169;color:#fff;}
        .result-box {background:#f0f7ff;padding:16px;border-radius:8px;margin-top:15px;white-space:pre-line;line-height:1.8;font-size:15px;}
        .tip {color:#718096;font-size:13px;}
    </style>
</head>
<body>
    <div class="card">
        <h2>一、基础投资参数</h2>
        <div class="grid-row">
            <div>
                <label>设备成本(元)</label>
                <input type="number" id="equipCost" value="2400000">
            </div>
            <div>
                <label>建筑工程费(元)</label>
                <input type="number" id="buildCost" value="100000">
            </div>
            <div>
                <label>运输/设计/预备费(合计)</label>
                <input type="number" id="otherCost" value="0">
            </div>
        </div>
        <div class="grid-two">
            <div>
                <label>贷款年利率(%)</label>
                <input type="number" id="loanRate" value="3" step="0.01">
            </div>
            <div>
                <label>贷款年限(年)</label>
                <input type="number" id="loanYear" value="10">
            </div>
        </div>
    </div>

    <div class="card">
        <h2>二、储能系统参数(2充2放)</h2>
        <div class="grid-row">
            <div>
                <label>系统功率kW</label>
                <input type="number" id="powerKw" value="250">
            </div>
            <div>
                <label>系统效率</label>
                <input type="number" id="eff" value="0.7" step="0.01">
            </div>
            <div>
                <label>年维护费用(元)</label>
                <input type="number" id="maintain" value="30000">
            </div>
        </div>
        <div class="grid-two">
            <div>
                <label>项目测算总年限(1-30)</label>
                <input type="number" id="projYear" value="25">
            </div>
            <div>
                <label>峰段每日充放次数</label>
                <input type="number" id="cycleNum" value="1">
            </div>
        </div>
    </div>

    <div class="card">
        <h2>三、分季节运行天数</h2>
        <div class="grid-row">
            <div>
                <label>2/3/4/5/9/10月(天)</label>
                <input type="number" id="day1" value="108">
            </div>
            <div>
                <label>6/7/8月(天)</label>
                <input type="number" id="day2" value="81.9">
            </div>
            <div>
                <label>11/12/1月(天)</label>
                <input type="number" id="day3" value="82.8">
            </div>
        </div>
    </div>

    <div class="card">
        <h2>四、分时电价 元/kWh</h2>
        <div class="grid-row">
            <div>
                <label>尖峰电价</label>
                <input type="number" id="pricePeak" value="0.999593" step="0.000001">
            </div>
            <div>
                <label>高峰电价</label>
                <input type="number" id="priceHigh" value="0.858435" step="0.000001">
            </div>
            <div>
                <label>平段电价</label>
                <input type="number" id="priceMid" value="0.567816" step="0.000001">
            </div>
        </div>
        <div style="width:33%">
            <label>谷段电价</label>
            <input type="number" id="priceValley" value="0.277197" step="0.000001">
        </div>
    </div>

    <div class="card">
        <button id="calcBtn">一键测算IRR与收益</button>
        <button id="resetBtn">重置为Excel基准参数</button>
        <button id="exportBtn">导出测算报告文本</button>
        <div class="result-box" id="resultBox">点击上方【一键测算IRR与收益】生成完整经济性结果</div>
    </div>

<script>
// IRR求解函数 牛顿迭代
function calcIRR(cashFlows, guess=0.05, maxIter=1000, tol=1e-8) {
    let r = guess;
    for(let i=0;i<maxIter;i++){
        let npv=0,deriv=0;
        for(let t=0;t<cashFlows.length;t++){
            let powT = Math.pow(1+r, t);
            npv += cashFlows[t]/powT;
            deriv -= t * cashFlows[t]/(powT*(1+r));
        }
        if(Math.abs(npv) < tol) return r;
        r = r - npv/deriv;
        if(r <= -1) r = -0.99;
    }
    return null;
}

// 等额本息总利息计算
function calcLoanTotalInterest(pv, rate, n) {
    let perRate = rate/100;
    let A = pv * perRate * Math.pow(1+perRate, n) / (Math.pow(1+perRate, n)-1);
    let totalPay = A * n;
    return totalPay - pv;
}

// 单时段收益计算（复刻Excel公式）
function calcSeasonIncome(days, cycle, power, eff, pPeak, pHigh, pMid, pValley){
    let chargeElec = cycle * power * days;
    let disElec = cycle * power * days * eff;
    // 收益1：谷充-平峰放
    let inc1 = (pMid + pHigh) * disElec * 2 - pValley * chargeElec * 4;
    // 收益2：平谷充-尖峰峰放（完全复刻Excel特殊公式）
    let inc2 = pHigh * inc1 * 4 - pValley * disElec * 2 - pMid * disElec * 2;
    let totalSeason = inc1 + inc2;
    return {inc1, inc2, totalSeason, chargeElec, disElec};
}

// 主测算逻辑
function runCalc(){
    // 读取输入
    const equipCost = Number(document.getElementById('equipCost').value);
    const buildCost = Number(document.getElementById('buildCost').value);
    const otherCost = Number(document.getElementById('otherCost').value);
    const loanRate = Number(document.getElementById('loanRate').value);
    const loanYear = Number(document.getElementById('loanYear').value);
    const powerKw = Number(document.getElementById('powerKw').value);
    const eff = Number(document.getElementById('eff').value);
    const maintain = Number(document.getElementById('maintain').value);
    const projYear = Number(document.getElementById('projYear').value);
    const cycleNum = Number(document.getElementById('cycleNum').value);
    const day1 = Number(document.getElementById('day1').value);
    const day2 = Number(document.getElementById('day2').value);
    const day3 = Number(document.getElementById('day3').value);
    const pPeak = Number(document.getElementById('pricePeak').value);
    const pHigh = Number(document.getElementById('priceHigh').value);
    const pMid = Number(document.getElementById('priceMid').value);
    const pValley = Number(document.getElementById('priceValley').value);

    // 1、总投资计算
    const loanPrincipal = equipCost;
    const loanInterest = calcLoanTotalInterest(loanPrincipal, loanRate, loanYear);
    const totalInvest = equipCost + buildCost + otherCost + loanInterest;

    // 2、三时段年度电费收益
    const s1 = calcSeasonIncome(day1, cycleNum, powerKw, eff, pPeak, pHigh, pMid, pValley);
    const s2 = calcSeasonIncome(day2, cycleNum, powerKw, eff, pPeak, pHigh, pMid, pValley);
    const s3 = calcSeasonIncome(day3, cycleNum, powerKw, eff, pPeak, pHigh, pMid, pValley);
    const annualElecIncome = s1.totalSeason + s2.totalSeason + s3.totalSeason;
    const annualNetCash = annualElecIncome - maintain;

    // 3、构建现金流：第0年初始投入，1~N年每年净现金流
    let cashFlow = [-totalInvest];
    for(let y=0;y<projYear;y++) cashFlow.push(annualNetCash);

    // 4、计算IRR、静态回收期
    const irr = calcIRR(cashFlow);
    let staticPayback = null;
    let accumulate = -totalInvest;
    for(let y=1;y<=projYear;y++){
        accumulate += annualNetCash;
        if(accumulate >= 0 && staticPayback === null){
            staticPayback = y - 1 + Math.abs(accumulate - annualNetCash) / annualNetCash;
        }
    }

    // 输出报告文本
    let report = `===== 储能峰谷套利经济性测算报告 =====
【基础投资】
设备成本：${equipCost.toLocaleString()} 元
建筑+其他：${(buildCost+otherCost).toLocaleString()} 元
10年期贷款总利息：${loanInterest.toFixed(2)} 元
项目总投资：${totalInvest.toFixed(2)} 元

【年度收益拆分】
春秋季年套利收益：${s1.totalSeason.toFixed(2)} 元
夏季年套利收益：${s2.totalSeason.toFixed(2)} 元
冬季年套利收益：${s3.totalSeason.toFixed(2)} 元
年度总电费收益：${annualElecIncome.toFixed(2)} 元
年运维成本：${maintain.toLocaleString()} 元
年度净现金流：${annualNetCash.toFixed(2)} 元

【核心财务指标】
静态投资回收期：${staticPayback ? staticPayback.toFixed(2)+" 年" : projYear+"年内无法回本"}
${projYear}年期内部收益率IRR：${irr ? (irr*100).toFixed(4)+" %" : "无有效解"}
`;
    document.getElementById('resultBox').innerText = report;
    window.reportCache = report;
}

// 重置按钮
document.getElementById('resetBtn').addEventListener('click',()=>{
    document.getElementById('equipCost').value = 2400000;
    document.getElementById('buildCost').value = 100000;
    document.getElementById('otherCost').value = 0;
    document.getElementById('loanRate').value = 3;
    document.getElementById('loanYear').value = 10;
    document.getElementById('powerKw').value = 250;
    document.getElementById('eff').value = 0.7;
    document.getElementById('maintain').value = 30000;
    document.getElementById('projYear').value = 25;
    document.getElementById('cycleNum').value = 1;
    document.getElementById('day1').value = 108;
    document.getElementById('day2').value = 81.9;
    document.getElementById('day3').value = 82.8;
    document.getElementById('pricePeak').value = 0.999593;
    document.getElementById('priceHigh').value = 0.858435;
    document.getElementById('priceMid').value = 0.567816;
    document.getElementById('priceValley').value = 0.277197;
    document.getElementById('resultBox').innerText = "已重置为Excel基准参数，点击测算生成结果";
});

// 导出文本
document.getElementById('exportBtn').addEventListener('click',()=>{
    if(!window.reportCache){
        alert("请先执行测算");
        return;
    }
    const blob = new Blob([window.reportCache],{type:'text/plain'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = "储能IRR测算报告.txt";
    a.click();
});

// 计算按钮
document.getElementById('calcBtn').addEventListener('click', runCalc);
</script>
</body>
</html>
