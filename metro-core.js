/* MetroYar route and UI engine */
(function(){
  const cache=new Map();
  let metroDataPromise;
  const ALL_STATIONS_CACHE=new Map();

  async function loadMetroData(){
    if(!metroDataPromise)metroDataPromise=fetch('metro.json',{cache:'no-cache'}).then(r=>{if(!r.ok)throw new Error('metro.json');return r.json();});
    return metroDataPromise;
  }
  function stationsForLine(line,mode){
    const main=[...line.main_route];
    if(mode==='branch' && line.branch_route){return [...main,...line.branch_route.filter(x=>!main.includes(x))];}
    return main;
  }
  function allStationsForParser(data){
    const a=[]; data.lines.forEach(l=>{l.main_route.forEach(s=>a.push(s));(l.branch_route||[]).forEach(s=>a.push(s));}); return [...new Set(a)];
  }
  function formatMinutes(n){
    const m=Math.max(0,Math.round(n)); return m===0?'همین الان':`${m} دقیقه دیگر`;
  }
  function excelServiceJourney(services,origin,destination){
    const out=[];
    for(const service of services){
      const oi=service.stationOrder.indexOf(origin), di=service.stationOrder.indexOf(destination);
      if(oi<0||di<0||oi>=di)continue;
      for(const tr of service.trips){
        const dep=tr.times[origin], arr=tr.times[destination];
        if(dep!=null&&arr!=null&&arr>=dep)out.push({depart:dep,arrival:arr,transfers:0,legs:[{from:origin,to:destination,depart:dep,arrive:arr,service}]});
      }
    }
    return out;
  }
  function nextLeg(services,from,to,after){
    let best=null;
    for(const service of services){
      const fi=service.stationOrder.indexOf(from), ti=service.stationOrder.indexOf(to);
      if(fi<0||ti<0||fi>=ti)continue;
      for(const tr of service.trips){const dep=tr.times[from],arr=tr.times[to];if(dep==null||arr==null||arr<dep||dep<after)continue; if(!best||dep<best.depart)best={depart:dep,arrival:arr,service};}
    }
    return best;
  }
  function transferJourneys(firstServices,secondServices,origin,hub,destination){
    const out=[];
    for(const service of firstServices){
      const oi=service.stationOrder.indexOf(origin), hi=service.stationOrder.indexOf(hub); if(oi<0||hi<0||oi>=hi)continue;
      for(const tr of service.trips){const dep=tr.times[origin],arr=tr.times[hub];if(dep==null||arr==null||arr<dep)continue; const second=nextLeg(secondServices,hub,destination,arr); if(!second)continue;
        out.push({depart:dep,arrival:second.arrival,transfers:1,legs:[{from:origin,to:hub,depart:dep,arrive:arr,service},{from:hub,to:destination,depart:second.depart,arrive:second.arrival,service:second.service}],hub});
      }
    }
    return out;
  }
  function dedupeSort(journeys){
    const map=new Map(); for(const j of journeys){const key=`${j.depart}|${j.arrival}|${j.transfers}|${j.hub||''}`;if(!map.has(key))map.set(key,j);}
    return [...map.values()].sort((a,b)=>a.depart-b.depart||a.arrival-b.arrival).slice(0,12);
  }
  function filterUpcoming(journeys,now){return dedupeSort(journeys.filter(j=>j.depart>=now));}

  async function buildPlanner(lineId,mode,day,origin,destination){
    const data=await loadMetroData(); const line=data.lines.find(x=>x.id===lineId); if(!line)throw new Error('خط پیدا نشد');
    const all=allStationsForParser(data); const sources=window.METRO_SCHEDULE_SOURCES[lineId]; if(!sources)throw new Error('منبع Excel خط پیدا نشد');
    const now=(new Date()).getHours()*60+(new Date()).getMinutes();
    const mainServices=await MetroExcel.loadSources(sources.main,all,day);
    let journeys=[];
    if(!sources.branch){
      if(lineId===5 && sources.branch){/* unreachable; kept for clarity */}
      else journeys=excelServiceJourney(mainServices,origin,destination);
    } else if(lineId===1 || lineId===4){
      const useBranch=mode==='branch';
      if(useBranch){
        const branchServices=await MetroExcel.loadSources(sources.branch,all,day);
        const hub=sources.branchHub;
        if(origin!==hub&&destination!==hub){
          const mainStations=new Set(line.main_route); const branchStations=new Set(line.branch_route||[]);
          if(mainStations.has(origin)&&mainStations.has(destination)) journeys=excelServiceJourney(mainServices,origin,destination);
          else if(branchStations.has(origin)&&branchStations.has(destination)) journeys=excelServiceJourney(branchServices,origin,destination);
          else if(mainStations.has(origin)&&branchStations.has(destination)) journeys=transferJourneys(mainServices,branchServices,origin,hub,destination);
          else if(branchStations.has(origin)&&mainStations.has(destination)) journeys=transferJourneys(branchServices,mainServices,origin,hub,destination);
        } else if(origin===hub&&destination!==hub){
          const branchStations=new Set(line.branch_route||[]); journeys=branchStations.has(destination)?excelServiceJourney(branchServices,origin,destination):excelServiceJourney(mainServices,origin,destination);
        } else if(destination===hub&&origin!==hub){
          const branchStations=new Set(line.branch_route||[]); journeys=branchStations.has(origin)?excelServiceJourney(branchServices,origin,destination):excelServiceJourney(mainServices,origin,destination);
        }
      } else journeys=excelServiceJourney(mainServices,origin,destination);
    } else if(lineId===5){
      const branchServices=await MetroExcel.loadSources(sources.branch,all,day); const hub='گلشهر';
      const mainSet=new Set(line.main_route); const branchRoute=['گلشهر','هشتگرد']; const branchSet=new Set(branchRoute);
      if(mainSet.has(origin)&&mainSet.has(destination)) journeys=excelServiceJourney(mainServices,origin,destination);
      else if(branchSet.has(origin)&&branchSet.has(destination)) journeys=excelServiceJourney(branchServices,origin,destination);
      else if(mainSet.has(origin)&&branchSet.has(destination)) journeys=transferJourneys(mainServices,branchServices,origin,hub,destination);
      else if(branchSet.has(origin)&&mainSet.has(destination)) journeys=transferJourneys(branchServices,mainServices,origin,hub,destination);
    }
    const upcoming=filterUpcoming(journeys,now);
    return {data,line,upcoming,now,mode};
  }

  function renderResult(box,planner,origin,destination){
    if(!planner.upcoming.length){box.innerHTML='<span class="error">دیر رسیدی، پایان فعالیت این سرویس 🙃</span>';return;}
    const items=planner.upcoming.slice(0,2); let html=`از <span class="highlight">${origin}</span> به <span class="highlight">${destination}</span><br><br>`;
    items.forEach((j,i)=>{
      const wait=j.depart-planner.now; const arrivalWait=j.arrival-planner.now;
      html+=`<div><b>قطار ${i+1}</b> — ${wait===0?'همین الان':`حدود <span class="highlight">${wait}</span> دقیقه دیگر`}</div>`;
      if(j.transfers){
        const h=j.legs[0].arrive-j.depart; const tw=j.legs[1].depart-j.legs[0].arrive;
        html+=`<div class="muted">تا <span class="highlight">${j.hub}</span>: حدود ${h} دقیقه</div>`;
        html+=`<div class="muted">تعویض قطار در <span class="highlight">${j.hub}</span> و حدود ${tw} دقیقه انتظار</div>`;
      }
      html+=`<div>رسیدن به مقصد: حدود <span class="highlight">${arrivalWait}</span> دقیقه دیگر</div>`;
      // Direction is expressed only with the service's terminal station.
      j.legs.forEach((leg,idx)=>{
        const dirEnd=leg.service.stationOrder[leg.service.stationOrder.length-1];
        html+=`<div class="muted">${j.legs.length>1?`حرکت ${idx+1}`:'حرکت'} به سمت <span class="highlight">${dirEnd}</span></div>`;
      });
      if(i===0&&items.length>1)html+='<hr style="border:0;border-top:1px solid rgba(255,255,255,.12);margin:12px 0">';
    });
    box.innerHTML=html;
  }

  window.MetroCore={loadMetroData,stationsForLine,buildPlanner,renderResult};
})();
