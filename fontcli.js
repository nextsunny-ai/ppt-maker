// CLI wrapper so the server can run font work in a child process (non-blocking).
//   node fontcli.js analyze <html>   -> {"fonts":[...]}
//   node fontcli.js install <name>   -> {"ok":true|false}
const fontkit = require('./fontkit');
const [,,cmd,arg] = process.argv;
(async()=>{
  try{
    if(cmd==='analyze'){
      const used = await fontkit.detectUsedFonts(arg);
      process.stdout.write(JSON.stringify({fonts:fontkit.analyzeFonts(used)}));
    } else if(cmd==='install'){
      process.stdout.write(JSON.stringify({ok:fontkit.installFont(arg)}));
    } else {
      process.stdout.write(JSON.stringify({error:'unknown cmd'}));
    }
  }catch(e){ process.stdout.write(JSON.stringify({error:String(e.message||e)})); }
})();
