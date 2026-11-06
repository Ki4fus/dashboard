import React from 'react';

const colors = ['#159c9f','#578fce','#7b78b7','#d5a35b','#6daba0','#a5819e','#8396ae','#b2b879'];
const amount = (value:number) => value.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
export function CurrencyBreakdown({currencies}:{currencies:{name:string;value:number}[]}) {
 const [selected,setSelected]=React.useState<string|null>(null);
 const total=currencies.reduce((sum,item)=>sum+item.value,0);
 const active=currencies.find(item=>item.name===selected);
 let offset=0;
 return <div className="currency-layout">
  <div className="currency-chart currency-donut">
   <svg viewBox="0 0 300 300" role="img" aria-label={`Currency breakdown: ${currencies.map(item=>`${item.name} ${amount(item.value/total*100)}%`).join(', ')}`}>
    <circle className="donut-track" cx="150" cy="150" r="113" fill="none" strokeWidth="30"/>
    {currencies.map((item,index)=>{
     const share=item.value/total;const start=offset;offset+=share;
     return <circle key={item.name} cx="150" cy="150" r="113" pathLength="100" fill="none" stroke={colors[index%colors.length]} strokeWidth={active?.name===item.name?36:30} strokeDasharray={`${share*100} ${100-share*100}`} strokeDashoffset={-start*100} transform="rotate(-90 150 150)" className={`donut-segment ${active && active.name!==item.name?'is-muted':''}`} onMouseEnter={()=>setSelected(item.name)} onMouseLeave={()=>setSelected(null)}>
      <title>{item.name}: {amount(item.value)} USD · {amount(share*100)}%</title>
     </circle>;
    })}
   </svg>
   <div className="donut-center" aria-hidden="true"><span>{active?.name ?? 'Total fixed amount'}</span><strong title={amount(active?.value ?? total)}>{(active?.value ?? total).toLocaleString('en-US',{notation:'compact',maximumFractionDigits:2})}</strong><small>USD · {active ? `${amount(active.value/total*100)}%` : `${currencies.length} ${currencies.length===1?'currency':'currencies'}`}</small></div>
  </div>
  <div className="currency-legend">{currencies.map((item,index)=><button type="button" className={`currency-entry ${active?.name===item.name?'is-selected':''}`} key={item.name} aria-pressed={active?.name===item.name} onClick={()=>setSelected(selected===item.name?null:item.name)} style={{'--currency-color':colors[index%colors.length]} as React.CSSProperties}>
   <span className="currency-entry-heading"><span className="currency-name"><span className="currency-dot"/>{item.name}</span><span className="currency-share">{amount(item.value/total*100)}%</span></span>
   <span className="numeric-value currency-amount">{amount(item.value)} <small>USD</small></span>
   <span className="currency-bar" aria-hidden="true"><span style={{width:`${item.value/total*100}%`}}/></span>
  </button>)}</div>
 </div>;
}
