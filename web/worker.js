/* No user files leave this worker. Only program assets are fetched from this origin. */
importScripts('./runtime/pyodide.js');
let py;
const ready=(async()=>{
  postMessage({progress:{stage:'正在加载计算组件，首次打开可能需要稍候'}});
  py=await loadPyodide({indexURL:new URL('./runtime/',self.location.href).href});
  await py.loadPackage(['sqlite3','hashlib','pydecimal']);
  py.FS.mkdirTree('/app');
  for(const name of ['engine.py','exporter.py','bridge.py','result_notes.py','template_headers.json','xlsxwriter.zip']){
    const r=await fetch('./'+name+'?v=3.1.0');if(!r.ok)throw Error('计算组件加载失败，请刷新后重试');
    py.FS.writeFile('/app/'+name,new Uint8Array(await r.arrayBuffer()));
  }
  py.runPython("import sys,zipfile\nsys.path.insert(0,'/app')\nzipfile.ZipFile('/app/xlsxwriter.zip').extractall('/app')\nimport bridge,json\n");
  py.globals.set('progress_js',data=>postMessage({progress:JSON.parse(data)}));
  py.runPython("def browser_progress(**kw):\n    progress_js(json.dumps(kw,ensure_ascii=False))");
})();
let queue=Promise.resolve();
self.onmessage=e=>{queue=queue.then(async()=>{
  const {id,url,payload,filename,buffer,download}=e.data;
  try{
    await ready;
    if(download){
      py.globals.set('kind_js',download);
      const path=py.runPython('bridge.download_path(kind_js)');
      const bytes=py.FS.readFile(path);postMessage({id,bytes:bytes.buffer,name:path.split('/').pop()},[bytes.buffer]);return;
    }
    if(buffer)py.FS.writeFile('/workspace/incoming.xlsx',new Uint8Array(buffer));
    py.globals.set('url_js',url);py.globals.set('payload_js',JSON.stringify(payload||{}));py.globals.set('filename_js',filename||'');
    const data=py.runPython('json.dumps(bridge.dispatch(url_js,json.loads(payload_js),filename_js,browser_progress),ensure_ascii=False)');
    postMessage({id,data:JSON.parse(data)});
  }catch(error){postMessage({id,error:String(error.message||error).split('\n').slice(-5).join('\n')});}
});};
