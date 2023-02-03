

export function readFile(file){
  return new Promise(resolve=> {
    let reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onload=ev=>{
      resolve(ev.target.result);
    }
  });
}


export function delay(interval=0){
  return new Promise(resolve=>{
    let timer = setTimeout(_ => {
      clearTimeout(timer);
      resolve();
    }, interval);
  });
}

export function getObjectUrl(file) {
  let url = null;
  if (window.createObjectURL != undefined) {
    // basic
    url = window.createObjectURL(file);
  } else if (window.webkitURL != undefined) {
    // webkit or chrome
    url = window.webkitURL.createObjectURL(file);
  } else if (window.URL != undefined) {
    // mozilla(firefox)
    url = window.URL.createObjectURL(file);
  }
  return url;
}


export let character={
  // name:{
  //   text:"姓名",
  //   type:"string"
  // },
  // phone:{
  //   text:"电话",
  //   type:"string"
  // }
}
