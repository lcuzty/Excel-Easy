const fs = require('fs')
const path = require('path')
const { exec } = require('child_process')
const os = require('os')
const {stat, promises} = require("fs");
const {resolve} = require("path");

class Base64String{
    constructor(n=0){
        this.data = "0"
        this.setByBase10(n)
    }

    _base64StringArrAdd0AtBegin(stringArr){
        let re = ['0']
        for(let i=0;i<stringArr.length;i++){
            re.push(stringArr[i])
        }
        return re
    }

    _base64StringCompare(baseString1,baseString2){

        if(string1.length<string2.length){
            return -1
        }
        if(string1.length>string2.length){
            return 1
        }
        for(let i=0;i<string1.length;i++){
            if(string1.slice(i,i+1).charCodeAt(0)<string2.slice(i,i+1).charCodeAt(0)){
                return -1
            }
            if(string1.slice(i,i+1).charCodeAt(0)>string2.slice(i,i+1).charCodeAt(0)){
                return 1
            }
        }
        return 0
    }

    _base64StringAdd1(string){
        string = string.split('')
        let flag = true
        let index = string.length - 1
        while (flag==true) {
            if(string[index]=='β'){
                string[index]='0'
                index-=1
                if(index==-1){
                    index=0
                    string = this._base64StringArrAdd0AtBegin(string)
                }
            }else{
                string[index] = String.fromCharCode(string[index].charCodeAt(0) + 1)
                if(string[index] == String.fromCharCode(58)){
                    string[index] = 'A'
                    flag = false
                }
                if(string[index] == String.fromCharCode(91)){
                    string[index] = 'a'
                    flag = false
                }
                if(string[index] == String.fromCharCode(123)){
                    string[index] = 'α'
                    flag = false
                }
                flag = false
            }
        }
        let re = ''
        for(let i=0;i<string.length;i++){
            re+=string[i]
        }
        return re
    }

    add1(){
        let string = this.data
        string = string.split('')
        let flag = true
        let index = string.length - 1
        while (flag==true) {
            if(string[index]=='β'){
                string[index]='0'
                index-=1
                if(index==-1){
                    index=0
                    string = this._base64StringArrAdd0AtBegin(string)
                }
            }else{
                string[index] = String.fromCharCode(string[index].charCodeAt(0) + 1)
                if(string[index] == String.fromCharCode(58)){
                    string[index] = 'A'
                    flag = false
                }
                if(string[index] == String.fromCharCode(91)){
                    string[index] = 'a'
                    flag = false
                }
                if(string[index] == String.fromCharCode(123)){
                    string[index] = 'α'
                    flag = false
                }
                flag = false
            }
        }
        let re = ''
        for(let i=0;i<string.length;i++){
            re+=string[i]
        }
        this.data = re
    }

    sub1(){
        let string = this.data
        let re = "0"
        let last = ""
        while (this._base64StringCompare(re,string)) {
            last = re
            re = this._base64StringAdd1(re)
        }
        this.data = re
    }

    setByBase10(n){
        this.data = "0"
        for(let i=0;i<n;i++){
            this.add1()
        }
    }

    setByBase64(string){
        string = decodeString(string)
        this.data = string
    }

    get(){
        return encodeString(this.data)
    }

    getBase10(){
        let re = 0
        let c = "0"
        while(c!=this.data){
            re+=1
            c = this._base64StringAdd1(c)
        }
        return re
    }
}




module.exports = {
    getColumnName(columnIndex) {
        let columnName = '';

        // 检查输入是否有效
        if (columnIndex < 1) {
            return 'Invalid column index';
        }

        // 转换列索引为列名
        while (columnIndex > 0) {
            columnIndex--; // 转换为从0开始的索引
            columnName = String.fromCharCode(65 + (columnIndex % 26)) + columnName; // 将索引转换为对应的字母并添加到列名前面
            columnIndex = Math.floor(columnIndex / 26); // 移除已处理的最后一位
        }

        return columnName;
    },
    normalizeBackslashes(inputString) {
        inputString = inputString.split('')
        if(inputString.length==0){
            return ''
        }
        let re = inputString[0]
        for(let i=1;i<inputString.length;i++){
            if(inputString[i-1]=='\\' && inputString[i]=='\\'){
                continue
            }
            re+=inputString[i]
        }
        return re
    },
    controlForm(id,op){
      let re = document.getElementsByClassName('form-bg')
      for(let i=0;i<re.length;i++){
        if(re[i].id==id){
          if(op){
            re[i].style.display = 'unset'
          }else{
            re[i].style.display = 'none'
          }
          return
        }
      }
    },
    getAllDrive(){
        return new Promise((resolve, reject) => {  
            exec('wmic logicaldisk get deviceid', (error, stdout, stderr) => {  
              if (error) {  
                reject(error);  
              } else {   
                let driveLetters = stdout.split('\n').slice(1).map(line => line.trim());  
                for(let i = driveLetters.length-1;i>=0;i--){
                    if(driveLetters[i]!='')continue
                    driveLetters = this.deleteArrElemByIndex(driveLetters,i)
                }
                resolve(driveLetters);  
              }  
            });  
          });  
    },
    deleteArrElemByIndex(arr,index){
        let re = []
        for(let i=0;i<arr.length;i++){
            if(i==index)continue
            re.push(arr[i])
        }
        return re
    },
    createFolder(folderName) {
      return new Promise((resolve, reject) => {
        // const folderPath = path.join(__dirname, folderName);
    
        fs.mkdir(folderName, { recursive: true }, (err) => {
          if (err) {
            reject(err);
          } else {
            return resolve();
          }
        });
      });
    },
    readFile(filePath) {
      return new Promise((resolve, reject) => {
        fs.readFile(filePath, 'utf8', (err, data) => {
          if (err) {
            reject(err)
          } else {
            return resolve(data)
          }
        })
      })
    },
    createFileWithContent(filePath, content) {
      return new Promise((resolve, reject) => {
        fs.writeFile(filePath, content, { flag: 'wx' }, (err) => {
          if (err) {
            reject(err);
          } else {
            return resolve();
          }
        });
      });
    },
    checkPathExists(fileOrFolderPath) {
      return new Promise((resolve) => {
        fs.access(fileOrFolderPath, fs.constants.F_OK, (err) => {
          if (err) {
            return resolve(false);
          } else {
            return resolve(true);
          }
        });
      });
    },
    updateFileContent(filePath, content) {
      return new Promise((resolve, reject) => {
        fs.unlink(filePath, (error) => {
          if (error && error.code !== 'ENOENT') {
            reject(error);
          } else {
            fs.writeFile(filePath, content, { flag: 'w+' }, (error) => {
              if (error) {
                reject(error);
              } else {
                resolve();
              }
            });
          }
        });
      });
    },
    getDirectoriesInFolder(folderPath) {
      return new Promise((resolve, reject) => {
        fs.readdir(folderPath, (err, files) => {
          if (err) {
            reject(err)
            return
          }
          const directories = files.filter(file =>
            fs.statSync(path.join(folderPath, file)).isDirectory()
          )
          return resolve(directories)
        })
      })
    },
    getFilesInFolder(folderPath) {
      return new Promise((resolve, reject) => {
        try {
          fs.readdir(folderPath, (err, files) => {
            if (err) {
              reject(err)
              return
            }
            const directories = files.filter(file => {
              try {
                return fs.statSync(path.join(folderPath, file)).isDirectory()==false
              } catch (error) {
                return false
              }
            }

            )
            return resolve(directories)
          })
        } catch (error) {
            console.log(error)
          return resolve([])
        }
      })
    },
    deleteFolder(folderName) {
      return new Promise(async (resolve, reject) => {
        const folderPath = folderName;
        if (!fs.existsSync(folderPath)) {
          return resolve()
        }
        await fs.readdirSync(folderPath).forEach(async (file) => {
          const filePath = path.join(folderPath, file);
          if (fs.lstatSync(filePath).isDirectory()) {
            await this.deleteFolder(filePath)
          } else {
            fs.unlinkSync(filePath)
          }
        })
        fs.rmdirSync(folderPath)
        return resolve()
      })
      
    },
    deleteFile(filePath) {
        return new Promise((resolve, reject) => {
            fs.unlink(filePath, (err) => {
                if (err) {
                    reject(err);
                } else {
                    resolve();
                }
            });
        });
    },
    renameFolder(folderPath, newFolderName) {
      return new Promise((resolve, reject) => {
        const parentPath = path.dirname(folderPath);
        const newPath = path.join(parentPath, newFolderName);
    
        fs.rename(folderPath, newPath, (err) => {
          if (err) {
            reject(err);
          } else {
            resolve(true);
          }
        });
      });
    },
    renameFile(filePath, newFileName) {
      return new Promise((resolve, reject) => {
        const fileDir = path.dirname(filePath);
        const newFilePath = path.join(fileDir, newFileName);
        fs.rename(filePath, newFilePath, (err) => {
          if (err) {
            resolve(false)
          } else {
            resolve(true)
          }
        });
      });
    },
    waitSeconds(n){
      return new Promise((resolve, reject) => {
        setTimeout(()=>{
          return resolve()
        },n*1000)
      })
    },
    async getFolderTree(parentFolderPath,showFile){

      function getDirectoriesInFolder(folderPath) {
        return new Promise((resolve, reject) => {
          try {
            fs.readdir(folderPath, (err, files) => {
              if (err) {
                reject(err)
                return
              }
              const directories = files.filter(file => {
                try {
                  return fs.statSync(path.join(folderPath, file)).isDirectory()
                } catch (error) {
                  return false
                }
              }
                
              )
              return resolve(directories)
            })
          } catch (error) {
            return resolve([])
          }
        })
      }

      function getFilesInFolder(folderPath) {
        return new Promise((resolve, reject) => {
          try {
            fs.readdir(folderPath, (err, files) => {
              if (err) {
                reject(err)
                return
              }
              const directories = files.filter(file => {
                try {
                  return fs.statSync(path.join(folderPath, file)).isDirectory()==false
                } catch (error) {
                  return false
                }
              }
                
              )
              return resolve(directories)
            })
          } catch (error) {
            return resolve([])
          }
        })
      }

      function getFileOrDirDateTime(filePath) {  
        return new Promise((resolve, reject) => {  
            fs.stat(filePath, (err, stats) => {  
                if (err) {  
                    reject(err);  
                } else {   
                    const dateTime = stats.isDirectory() ? stats.birthtime : stats.mtime;  
                    resolve(dateTime);  
                }  
            });  
        });  
      } 

      async function get(parentFolderPath){
        let re = []
        let folders = await getDirectoriesInFolder(parentFolderPath)
        let files = await getFilesInFolder(parentFolderPath)
        for(let i=0;i<folders.length;i++){
          re.push({
            name:folders[i],
            children:await get(parentFolderPath + "\\" + folders[i]),
            time:await getFileOrDirDateTime(parentFolderPath + "\\" + folders[i])
          })
        }
        if(showFile){
          for(let i=0;i<files.length;i++){
            re.push({
              name:files[i],
              time:await getFileOrDirDateTime(parentFolderPath + "\\" + files[i])
            })
          }
        }
        return re
      }

      return get(parentFolderPath)
    },
    isValidFolderName(folderName) {  
      const invalidChars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|'];   
      for (const char of invalidChars) {  
        if (folderName.includes(char)) {  
          return false;  
        }  
      }  
      if (folderName.trim() !== folderName) {  
        return false;  
      }   
      const reservedNames = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'];  
        
      if (reservedNames.includes(folderName.toUpperCase())) {  
        return false;  
      }  
      if (folderName.length > 255) {  
        return false;  
      }  
      return true;  
    },
    isValidFileName(fileName) {  
      const invalidChars = /[\<>:"\/\\|?*]/;  
      if (invalidChars.test(fileName)) {  
          return false;  
      }  
      if (fileName.startsWith(' ') || fileName.endsWith(' ')) {  
          return false;  
      }  
      if (fileName.length > 255) {  
          return false;  
      }  
      const reservedNames = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'];  
      if (reservedNames.includes(fileName.toUpperCase())) {  
          return false;  
      }   
      return true;  
    },
    getFileOrDirDateTime(filePath) {  
      return new Promise((resolve, reject) => {  
          fs.stat(filePath, (err, stats) => {  
              if (err) {  
                  reject(err);  
              } else {   
                  const dateTime = stats.isDirectory() ? stats.birthtime : stats.mtime;  
                  resolve(dateTime);  
              }  
          });  
      });  
    },
    formatDate(date) {  
      const options = {  
        year: 'numeric',  
        month: '2-digit',  
        day: '2-digit',  
        hour: '2-digit',  
        minute: '2-digit',  
        second: '2-digit',  
      };  
      
      return new Intl.DateTimeFormat('en-US', options).format(date);  
    },
    async moveFile(src, dest) {
        try {
            await fs.rename(src, dest,(err)=>{
                console.log(err)
            });
            console.log(`File moved successfully from ${src} to ${dest}`);
        } catch (err) {
            if (err.code === 'EXDEV') {
                // `rename` failed because src and dest are on different filesystems
                const data = await fs.readFile(src);
                await fs.writeFile(dest, data);
                await fs.unlink(src);
                console.log(`File moved successfully from ${src} to ${dest} using copy and delete`);
            } else {
                // Some other error, rethrow
                throw err;
            }
        }
    },
    getFileNameFromFilePath(filePath){
        filePath = filePath.split('\\')
        return filePath[filePath.length-1]
    },
    getFileNameFormFileWithoutExt(filePath){
        filePath = this.getFileNameFromFilePath(filePath)
        filePath = filePath.split('.')
        let re = ''
        for(let i=0;i<filePath.length-1;i++){
            if(i){
                re+='.'
            }
            re+=filePath[i]
        }
        return re
    },
    isValidWorksheetName(name) {
        name = name.split('')
        if(name.length<1 || name.length>31){
            return false
        }
        let check = [':','\\',"/",'?','*','[',']']
        for(let i=0;i<name.length;i++){
            for(let i1=0;i1<check.length;i1++){
                if(name[i]==check[i1]){
                    return false
                }
            }
        }
        return true;
    },
    getFormattedLastModifiedTime(filePath) {
        try {
            const stats = fs.statSync(filePath);
            const lastModifiedTime = new Date(stats.mtime);
            const year = lastModifiedTime.getFullYear();
            const month = String(lastModifiedTime.getMonth() + 1).padStart(2, '0'); // 月份从0开始，所以+1，并使用padStart补0
            const day = String(lastModifiedTime.getDate()).padStart(2, '0'); // 使用padStart补0
            const hours = String(lastModifiedTime.getHours()).padStart(2, '0'); // 使用padStart补0
            const minutes = String(lastModifiedTime.getMinutes()).padStart(2, '0'); // 使用padStart补0
            const seconds = String(lastModifiedTime.getSeconds()).padStart(2, '0'); // 使用padStart补0

            return `${year}年${month}月${day}日 ${hours}:${minutes}:${seconds}`;
        } catch (error) {
            console.error(`Error getting last modified time for ${filePath}:`, error);
            return null;
        }
    },
    Base64String,
    arrInsert(arr,index,item){
        let re = []
        for(let i=0;i<arr.length;i++){
            if(i==index)re.push(item)
            re.push(arr[i])
        }
        if(index==arr.length){
            re.push(item)
        }
        return re
    },
    chooseFolder() {
        return new Promise((resolve, reject) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.multiple = true;
            input.accept = '.xlsx, .xls';
            input.webkitdirectory = true;

            input.onchange = (event) => {
                const files = event.target.files;
                if (files && files.length > 0) {
                    const fileArray = [];
                    for (let i = 0; i < files.length; i++) {
                        const file = files[i];
                        const fileName = file.name;
                        const fileExtension = fileName.split('.').pop().toLowerCase();
                        if (fileExtension === 'xlsx' || fileExtension === 'xls') {
                            const fileReader = new FileReader();
                            fileReader.onload = (event) => {
                                const fileContent = event.target.result;
                                const fileData = {
                                    path: URL.createObjectURL(file),
                                    name: fileName
                                };
                                fileArray.push(fileData);
                                if (fileArray.length === files.length) {
                                    console.log(fileArray)
                                    resolve(fileArray);
                                }
                            };
                            fileReader.readAsArrayBuffer(file);
                        }
                    }
                } else {
                    reject(false);
                }
            };

            input.click();
        });
    },
    selectFile() {
        return new Promise((resolve, reject) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.xlsx, .xls';

            input.addEventListener('change', (event) => {
                const file = event.target.files[0];
                if (file) {
                    const path = URL.createObjectURL(file);
                    resolve({ path, fileName: file.name });
                } else {
                    resolve(false);
                }
            });

            input.click();
        });
    },
    removeFileNameExt(fileName){
        fileName = fileName.split('.')
        let re = ''
        for(let i=0;i<fileName.length-1;i++){
            if(i){
                re+='.'
            }
            re+=fileName[i]
        }
        return re
    },
    sortArray(arr1, arr2, isNum) {
        const sortedArr2 = [...arr2]; // 创建一个副本，避免修改原数组

        sortedArr2.sort((a, b) => {
            if(isNum){
                return arr1[a['KpAF0']-1] - arr1[b['KpAF0']-1];
            }else{
                if(arr1[a['KpAF0']-1]>arr1[b['KpAF0']-1])return 1
                if(arr1[a['KpAF0']-1]==arr1[b['KpAF0']-1])return 0
                if(arr1[a['KpAF0']-1]<arr1[b['KpAF0']-1])return -1
            }
        });

        return sortedArr2;
    },
    getNumberFromString(string){
        try{
            string = JSON.parse(string)
            if(string<0){
                throw new Error('')
            }
            if(parseInt(string)!=string){
                throw new Error('')
            }
            return string
        }catch (e) {
            return -1
        }
    },
}

//

function getExcelFilesFromFolder() {
    return new Promise((resolve, reject) => {
        // 检查浏览器是否支持 showOpenPicker
        if (!('showOpenPicker' in showDirectoryPicker)) {
            reject(new Error('Browser does not support showOpenPicker'));
            return;
        }

        // 打开文件夹选择器
        showDirectoryPicker()
            .then(handleDirectory)
            .then(resolve)
            .catch(reject);
    });
}

async function handleDirectory(directoryHandle) {
    const entries = await directoryHandle.getEntries();
    const excelFiles = [];

    for (const entry of entries) {
        if (entry.kind === 'file') {
            const file = await entry.getFile();
            const name = file.name;
            const path = file.path;

            // 检查文件扩展名是否为 .xlsx 或 .xls
            if (/\.(xlsx|xls)$/i.test(name)) {
                const json = { path, name };
                excelFiles.push(json);
            }
        }
    }

    return excelFiles;
}

function encodeString(string){
    function getNumberString(n,x){
        n = JSON.stringify(n)
        for(let i = n.length;i<x;i++){
            n = '0' + n
        }
        return n
    }
    string = string.split('')
    let string1 = ''
    for(let i=0;i<string.length;i++){
        string1 += JSON.stringify(JSON.stringify(string[i].charCodeAt(0)).length) + JSON.stringify(string[i].charCodeAt(0))
    }
    string1 = string1.split('')
    function isLetter(code){
        if(code>=65 && code<=90){
            return 1
        }
        if(code>=97 && code<=122){
            return 1
        }
        if(code>90 && code<97){
            return 2
        }
        if(code<65){
            return 0
        }
        return 2
    }
    for(let i=0;i<string1.length;i++){
        if(string1[i]=='0'){
            continue
        }
        let c = i
        let ss = string1[c]
        while (true) {
            let cd = isLetter(parseInt(ss))
            if(cd==2){
                break
            }
            if(cd==1){
                string1[c] = String.fromCharCode(parseInt(ss))
                for(let i1=i;i1<c;i1++){
                    string1[i1] = '-'
                }
                i = c
                break
            }
            c++
            if(c==string1.length){
                break
            }
            ss+=string1[c]
        }
    }
    for(let i=0;i<string1.length-1;i++){
        if(string1[i]=='2' && string1[i+1]=='2'){
            string1[i]='?'
            string1[i+1] = '-'
        }
    }
    let string2 = []
    for(let i=0;i<string1.length;i++){
        if(string1[i]=='-'){
            continue
        }
        string2.push(JSON.stringify((string1[i].charCodeAt(0) - string1[i].charCodeAt(0)%52)/52) + getNumberString(string1[i].charCodeAt(0)%52,2))
    }
    let string3 = ''
    for(let i = string2.length-1;i>=0;i--){
        string3 += string2[i]
    }
    string3 = string3.split('')
    let string4 = []
    for(let i=0;i<string3.length;i+=2){
        let ch = string3[i]
        if(i+1!=string3.length){
            ch+=string3[i+1]
        }
        if(parseInt(ch)<52 && ch.length==2){
            if(parseInt(ch)<26){
                string4.push(String.fromCharCode(parseInt(ch) + 65))
            }else{
                string4.push(String.fromCharCode(parseInt(ch) + 97 - 26))
            }
        }else{
            string4.push(ch)
        }
    }
    let string5 = ''
    for(let i=0;i<string4.length;i++){
        string5+=string4[i]
    }
    return string5
}
function decodeString(string){
    function getNumberString(n,x){
        n = JSON.stringify(n)
        for(let i = n.length;i<x;i++){
            n = '0' + n
        }
        return n
    }
    string = string.split('')
    for(let i=0;i<string.length;i++){
        if(string[i].charCodeAt(0)>=48 && string[i].charCodeAt(0)<=56){
            continue
        }
        if(string[i].charCodeAt(0)<=90){
            string[i] = getNumberString(string[i].charCodeAt(0) - 65,2)
        }else{
            string[i] = getNumberString(string[i].charCodeAt(0) - 97 + 26,2)
        }
    }
    let string2 = ''
    for(let i=0;i<string.length;i++){
        if(string[i]=='-8'){
            string2+='9'
            continue
        }
        string2+=string[i]
    }
    string2 = string2.split('')
    let string3 = []
    for(let i=0;i<string2.length;i+=3){
        string3.push(getNumberString(parseInt(string2[i])*52 + parseInt(string2[i+1] + string2[i+2]),3))
    }
    let string4 = []
    for(let i=string3.length-1;i>=0;i--){
        string4.push(String.fromCharCode(parseInt(string3[i])))
    }
    let string5 = ''
    for(let i=0;i<string4.length;i++){
        if(string4[i]=='?'){
            string5+='22'
            continue
        }
        if((string4[i].charCodeAt(0)>=48 && string4[i].charCodeAt(0)<=57)==false){
            string5+=JSON.stringify(string4[i].charCodeAt(0))
            continue
        }
        string5+=string4[i]
    }
    string5 = string5.split('')
    let string6 = ''
    for(let i=0;i<string5.length;i++){
        let n = parseInt(string5[i])
        let s = ''
        for(let i1=0;i1<n;i1++){
            s+=string5[i+1+i1]
        }
        string6+=String.fromCharCode(parseInt(s))
        i+=n
    }
    return string6
}