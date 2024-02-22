<template>
  <div @mousedown="checkMouseDownEvent" @mousemove="getMousePosition" style="width: 100vw;height: 100vh;position: relative;overflow: hidden;">

    <div v-if="data.setting.readData!=undefined" style="width: 100%;height: 48px;position: relative;z-index: 1000"
      :style="{
        color:data.themes[data.setting.readData.theme].titleBarColor,
        background:data.themes[data.setting.readData.theme].titleBarBackground,
      }"
    >
      <div v-if="data.showStartPage==0" style="display: inline-block;width: 48px;height: 48px;padding-left: 24px;padding-top: 24px;" class="button" @click="showStartPage" title="欢迎界面">
        <svg style="display: inline-block;transform: translate(-50%,-50%);" width="20" height="20" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M43 38a1 1 0 011 1v2a1 1 0 01-1 1H5a1 1 0 01-1-1v-2a1 1 0 011-1h38zm0-16a1 1 0 011 1v2a1 1 0 01-1 1H5a1 1 0 01-1-1v-2a1 1 0 011-1h38zm0-16a1 1 0 011 1v2a1 1 0 01-1 1H5a1 1 0 01-1-1V7a1 1 0 011-1h38z" fill="currentColor"/></svg>
      </div>
      <div style="display: inline-block;height: 48px;width: calc(100vw - 48px * 4);-webkit-app-region:drag;padding-left: 0px;position: absolute;line-height: 48px;font-size: 14px;text-align: center;padding-left: 96px;left: 48px;"
           :style="{
        color:data.themes[data.setting.readData.theme].titleBarColor,
      }">
        Excel Easy
      </div>
      <div style="float: right;position: relative;z-index: 5000000000;">
        <div title="最小化" style="display: inline-block;width: 48px;height: 48px;padding-left: 24px;padding-top: 24px;background-size: contain;background-repeat: no-repeat;background-position: center;" :class="'button' + data.themes[data.setting.readData.theme].controlButtonColor + ' min' + data.themes[data.setting.readData.theme].controlButtonColor" @click="controlWindow('min')">

        </div>
        <div :title="data.cwm?'还原':'最大化'" style="display: inline-block;width: 48px;height: 48px;padding-left: 24px;padding-top: 24px;background-size: contain;background-repeat: no-repeat;background-position: center;" class="button" @click="controlWindow('max')" :class="data.cwm?[('button' + data.themes[data.setting.readData.theme].controlButtonColor + ' restore' + data.themes[data.setting.readData.theme].controlButtonColor)]:[('button' + data.themes[data.setting.readData.theme].controlButtonColor + ' max' + data.themes[data.setting.readData.theme].controlButtonColor)]">
        
        </div>
        <div title="关闭" style="display: inline-block;width: 48px;height: 48px;padding-left: 24px;padding-top: 24px;background-size: contain;background-repeat: no-repeat;background-position: center;" class="button close-button close" @click="controlWindow('close')" :class="'button' + data.themes[data.setting.readData.theme].controlButtonColor + ' close-button close' + data.themes[data.setting.readData.theme].controlButtonColor">
        
        </div>
      </div>
    </div>

    <div id="rightMenu" class="form-bg" style="display: none;padding-left: 0;padding-top: 0;background-color: transparent;">
      <div @mousedown="hideRightMenu" style="position: absolute;left: 0;top: 0;width: 100%;height: 100%;">

      </div>
      <div @click="hideRightMenu" style="position: absolute;"
          :style="{
            left:this.data.rightMenu.left.toString() + 'px',
            top:this.data.rightMenu.top.toString() + 'px',
          }">

        <div v-if="'fileListFile'==data.rightMenu.name" class="rightMenu">
          <div class="rightMenuButton" @click="fileList_delete">删除</div>
          <div class="rightMenuButton" @click="data.new.newFileName = getFileNameFormFileWithoutExt(data.rightMenu.data);controlForm('rename',1)">重命名</div>
          <div class="rightMenuButton" @click="fileList_move">移动</div>
        </div>

        <div v-if="'fileListDir'==data.rightMenu.name" class="rightMenu">
          <div class="rightMenuButton" @click="fileList_delete">删除</div>
          <div class="rightMenuButton" @click="data.new.newFileName = getFileNameFormFileWithoutExt(data.rightMenu.data);controlForm('rename',1)">重命名</div>
        </div>

        <div v-if="'fileBarMenu'==data.rightMenu.name" class="rightMenu">
          <div class="rightMenuButton" @click="file_saveAll">全部保存</div>
          <div class="rightMenuButton" @click="file_closeAll">全部关闭</div>
          <div class="rightMenuButton" @click="file_saveAllClose">全部保存并关闭</div>
          <div class="rightMenuButton" @click="()=>{
            newButtonClick()
          }">新建表格文件</div>
        </div>

        <div v-if="'fileBarFileItemMenu'==data.rightMenu.name" class="rightMenu">
          <div class="rightMenuButton" @click="file_saveSingle">保存</div>
          <div class="rightMenuButton" @click="file_closeSingle">关闭</div>
          <div class="rightMenuButton" @click="async ()=>{
            await file_saveSingle()
            await file_closeSingle()
          }">保存并关闭</div>
        </div>

        <div v-if="'sheetBar'==data.rightMenu.name" class="rightMenu">
          <div class="rightMenuButton" @click="data.new.newFileName = '';controlForm('newSheet',1)">新建表格</div>
        </div>

        <div v-if="'sheetBarItem'==data.rightMenu.name" class="rightMenu">
          <div class="rightMenuButton" @click="()=>{
            showInputForm('重命名表格','新表格名',async ()=>{
              if(data.input.text==''){
                setWarningFormTitleAndContentAndShowForm('无法重命名','请输入表格名称。',false,()=>{})
                return
              }

              if(tool.isValidWorksheetName(data.input.text)==false){
                setWarningFormTitleAndContentAndShowForm('无法重命名','表格名称不合法。',false,()=>{})
                return
              }
              let c = 0
              for(let i=0;i<data.files[data.currentFile.index-1].data.sheets.length;i++){
                if(data.files[data.currentFile.index-1].data.sheets[i].name==data.input.text){
                  c+=1
                }
              }
              if(c>0){
                setWarningFormTitleAndContentAndShowForm('无法重命名','表格名称已存在或表格名称未变。',false,()=>{})
                return
              }
              for(let i=0;i<data.files[data.currentFile.index-1].data.sheets.length;i++){
                if(data.rightMenu.data.name==data.files[data.currentFile.index-1].data.sheets[i].name){
                  data.files[data.currentFile.index-1].data.sheets[i].name = data.input.text
                  break
                }
              }
              data.files[data.currentFile.index-1].system.unsave = true
              data.files[data.currentFile.index-1].system.currentSheetName = data.input.text
              controlForm('input',false)
            },data.rightMenu.data.name)
          }">重命名</div>
          <div class="rightMenuButton" @click="()=>{
            if(data.files[data.currentFile.index-1].data.sheets.length==1){
              setWarningFormTitleAndContentAndShowForm('无法删除表格','表格文件中至少有一个表格。',false,()=>{})
              return
            }
            setWarningFormTitleAndContentAndShowForm('删除表格','表格删除后不能恢复，是否继续？',true,()=>{
            for(let i=0;i<data.files[data.currentFile.index-1].data.sheets.length;i++){
              if(data.files[data.currentFile.index-1].data.sheets[i].name==data.rightMenu.data.name){
                if(data.files[data.currentFile.index-1].system.currentSheetIndex==i+1){
                  data.files[data.currentFile.index-1].system.currentSheetIndex = 1
                  data.files[data.currentFile.index-1].system.currentSheetName = data.files[data.currentFile.index-1].data.sheets[0].name
                }
                data.files[data.currentFile.index-1].data.sheets = tool.deleteArrElemByIndex(data.files[data.currentFile.index-1].data.sheets,i)
                data.files[data.currentFile.index-1].system.unsave = true
                break
              }
            }
          })}">删除</div>
          <div class="rightMenuButton" @click="()=>{
            let x = 0
            for(let i=0;i<data.files[data.currentFile.index-1].data.sheets.length;i++){
              if(data.files[data.currentFile.index-1].data.sheets[i].name==data.rightMenu.data.name){
                x = i
                break
              }
            }
            data.files[data.currentFile.index-1].system.unsave = true
            data.files[data.currentFile.index-1].data.sheets = moveArrElemItem(data.files[data.currentFile.index-1].data.sheets,x,true)
          }">左移</div>
          <div class="rightMenuButton" @click="()=>{
            let x = 0
            for(let i=0;i<data.files[data.currentFile.index-1].data.sheets.length;i++){
              if(data.files[data.currentFile.index-1].data.sheets[i].name==data.rightMenu.data.name){
                x = i
                break
              }
            }
            data.files[data.currentFile.index-1].system.unsave = true
            data.files[data.currentFile.index-1].data.sheets = moveArrElemItem(data.files[data.currentFile.index-1].data.sheets,x,false)
          }">右移</div>
        </div>

        <div v-if="'tableCell'==data.rightMenu.name" class="rightMenu">
          <div class="rightMenuLabel">单元格{{ '（位于第' + (data.rightMenu.data.columnIndex+1).toString() + '列,第' + (data.rightMenu.data.dataIndex + 1).toString() + '行）' }}</div>
          <div class="rightMenuButton" @click="()=>{
            let trs = document.getElementById('mainTable').getElementsByTagName('tr')
            let re = []
            for(let i=1;i<trs.length;i++){
              re.push(trs[i].getElementsByTagName('td'))
            }
            clipboard.writeText(re[data.rightMenu.data.dataIndex][data.rightMenu.data.columnIndex].innerText.toString())
          }">复制</div>
<!--          <div class="rightMenuButton" @click="()=>{-->
<!--            -->
<!--          }">-->
<!--            加粗-->
<!--          </div>-->
          <div class="rightMenuDivider"></div>
          <div class="rightMenuLabel">{{ '第' + (data.rightMenu.data.dataIndex + 1).toString() + '行' }}</div>
          <div class="rightMenuButton" @click="()=>{
              data.addRow.inputValues = addRow_getCurrentSheetCols()
              data.addRow.inputValues[0].value = data.rightMenu.data.dataIndex+1
              controlForm('addRow',1)
          }">上方插入</div>
          <div class="rightMenuButton" @click="()=>{
              data.addRow.inputValues = addRow_getCurrentSheetCols()
              data.addRow.inputValues[0].value = data.rightMenu.data.dataIndex+2
              controlForm('addRow',1)
          }">下方插入</div>
          <div class="rightMenuButton" @click="()=>{
            let trs = document.getElementById('mainTable').getElementsByTagName('tr')
            let re = []
            for(let i=1;i<trs.length;i++){
              re.push(trs[i].getElementsByTagName('td'))
            }
            let text = ''
            for(let i=0;i<trs[0].getElementsByTagName('td').length;i++){
              if(text!=''){
                text+='\n'
              }
              text+=trs[0].getElementsByTagName('td')[i].innerText + '：' +  re[data.rightMenu.data.dataIndex][i].innerText.toString()
            }
            clipboard.writeText(text)
          }">复制</div>
          <div class="rightMenuButton" @click="()=>{
            let cc = Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)
            let cd = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[data.rightMenu.data.dataIndex]
            let re = []
            for(let i=1;i<cc.length;i++){
              if(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].type!='text' && data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].type!='number' && data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].type!='date')continue
              re.push({
                keyName: cc[i],
                value: cd[cc[i]],
                type:data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].type,
                name:data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].name
              })
            }
            data.addRow.inputValues = re
            data.addRow.isEdit = true
            controlForm('addRow',1)
          }">编辑</div>
          <div class="rightMenuButton" @click="()=>{
            if(data.rightMenu.data.dataIndex==0){
              setWarningFormTitleAndContentAndShowForm('无法移动','已是第一行。',false,()=>{})
              return
            }
            operationStackAppend()
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data = moveArrElemItem(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data,data.rightMenu.data.dataIndex,true)
            for(let i=0;i<data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length;i++){
              data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i]['KpAF0'] = i+1
            }
            setCurrentFileUnsave()
          }">上移</div>
          <div class="rightMenuButton" @click="()=>{
            if(data.rightMenu.data.dataIndex==data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length-1){
              setWarningFormTitleAndContentAndShowForm('无法移动','已是最后一行。',false,()=>{})
              return
            }
            operationStackAppend()
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data = moveArrElemItem(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data,data.rightMenu.data.dataIndex,false)
            for(let i=0;i<data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length;i++){
              data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i]['KpAF0'] = i+1
            }
            setCurrentFileUnsave()
          }">下移</div>
          <div class="rightMenuButton" @click="()=>{
            showInputForm('移动','请输入移动到的序号。',()=>{
              if(checkIsNumber(data.input.text)==false){
                setWarningFormTitleAndContentAndShowForm('无法移动','序号应为正整数。',false,()=>{})
                return
              }
              let index = JSON.parse(data.input.text)
              if(parseInt(index)!=index){
                setWarningFormTitleAndContentAndShowForm('无法移动','序号应为正整数。',false,()=>{})
                return
              }
              if(index<1 || index>data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length){
                setWarningFormTitleAndContentAndShowForm('无法移动','序号不在记录数范围内，序号最大不能超过' + (data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length).toString() + '。',false,()=>{})
                return
              }
              let re = JSON.parse(JSON.stringify(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data))
              let re1 = []
              for(let i=0;i<re.length;i++){
                if(i==data.rightMenu.data.dataIndex){
                  continue
                }
                if(index-1>data.rightMenu.data.dataIndex){
                re1.push(re[i])
                  if(i==index-1){
                  re1.push(re[data.rightMenu.data.dataIndex])
                }
                }else{
                  if(i==index-1){
                  re1.push(re[data.rightMenu.data.dataIndex])
                }
                re1.push(re[i])
                }

              }
              operationStackAppend()
              data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data = re1
              for(let i=0;i<data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length;i++){
              data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i]['KpAF0'] = i+1
              }
              setCurrentFileUnsave()
              controlForm('input',false)
            },'')
          }">移动</div>
          <div class="rightMenuButton" @click="()=>{
            operationStackAppend()
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data = deleteArrElemByIndex(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data,data.rightMenu.data.dataIndex)
            for(let i=0;i<data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length;i++){
              data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i]['KpAF0'] = i+1
            }
            setCurrentFileUnsave()
          }">删除</div>
        </div>

        <div v-if="'tableTitleCell'==data.rightMenu.name" class="rightMenu">
          <div class="rightMenuLabel">{{ '第' + (data.rightMenu.data.columnIndex+1).toString() + '列' }}</div>
          <div class="rightMenuButton" @click="()=>{
            data.editCol.name = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].name
            data.editCol.sumType = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].sumType
            data.editCol.type = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].type
            data.editCol.minWidth = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].minWidth
            if(data.editCol.minWidth==undefined){
              data.editCol.minWidth = ''
            }
            data.editCol.keyName = data.rightMenu.data.columnKey
            data.editCol.isIndexCol = false
            controlForm('editCol',1)
          }">编辑</div>
          <div class="rightMenuButton" @click="()=>{
            if(data.rightMenu.data.columnIndex==1){
              setWarningFormTitleAndContentAndShowForm('无法左移','序号列不可移动，已在最左侧。',false,()=>{})
              return
            }
            let cc = Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)
            let co = Object.values(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)
            operationStackAppend()
            cc = moveArrElemItem(cc,data.rightMenu.data.columnIndex,1)
            co = moveArrElemItem(co,data.rightMenu.data.columnIndex,1)
            let re = {}
            for(let i=0;i<cc.length;i++){
              re[cc[i]] = co[i]
            }
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns = re
            setCurrentFileUnsave()
            refreshTopAndBottom()
          }">左移</div>
          <div class="rightMenuButton" @click="()=>{
            if(data.rightMenu.data.columnIndex==data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length){
              setWarningFormTitleAndContentAndShowForm('无法右移','已在最右侧。',false,()=>{})
              return
            }
            let cc = Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)
            let co = Object.values(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)
            operationStackAppend()
            cc = moveArrElemItem(cc,data.rightMenu.data.columnIndex,0)
            co = moveArrElemItem(co,data.rightMenu.data.columnIndex,0)
            let re = {}
            for(let i=0;i<cc.length;i++){
              re[cc[i]] = co[i]
            }
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns = re
            setCurrentFileUnsave()
            refreshTopAndBottom()
          }">右移</div>
          <div class="rightMenuButton" @click="()=>{
            operationStackAppend()
            delete data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey]
            for(let i=0;i<data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length;i++){
              delete data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i][data.rightMenu.data.columnKey]
            }
            if(Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns).length==1){
              data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data = []
            }
            setCurrentFileUnsave()
            refreshTopAndBottom()
          }">删除</div>
        </div>

        <div v-if="'tableIndexTitleCell'==data.rightMenu.name" class="rightMenu">
          <div class="rightMenuLabel">序号列</div>
          <div class="rightMenuButton" @click="()=>{
            data.editCol.name = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].name
            data.editCol.sumType = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].sumType
            data.editCol.type = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].type
            data.editCol.keyName = data.rightMenu.data.columnKey
            data.editCol.minWidth = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].minWidth
            if(data.editCol.minWidth==undefined){
              data.editCol.minWidth = ''
            }
            data.editCol.isIndexCol = true
            controlForm('editCol',1)
          }">编辑</div>
        </div>

      </div>
    </div>

    <div id="setting" class="form-bg" style="display: none;">
      <div class="form-window">
        <div class="form-title">
          设置
          <svg v-if="data.appStarted" @click="controlForm('setting',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content">
          <p>存储位置</p>
          <select v-model="data.setting.form.saveDrive" style="width: 200px;margin-top: 10px;">
            <option v-for="item in data.setting.drives" :value="item">{{ item.slice(0,item.length-1) + '盘' }}</option>
          </select>
          <div class="tip">
            <p class="tip-title">提示</p>
            <p>修改存储位置将影响其他盘的数据可见性。</p>  
            <p>建议将存储位置设置为本地磁盘。</p>  
          </div>
          <p style="margin-top: 10px">关闭所有文件后</p>
          <select v-model="data.setting.form.closeAllFiles" style="width: 200px;margin-top: 10px;">
            <option value="exit">退出程序</option>
            <option value="showStart">打开欢迎界面</option>
          </select>
          <p style="margin-top: 10px">主题</p>
          <select v-model="data.setting.form.theme" style="width: 200px;margin-top: 10px;">
            <option v-for="item in Object.keys(data.themes)" :value="item">{{ data.themes[item].name }}</option>
          </select>
          <p style="margin-top: 10px">在欢迎界面上显示每日背景图</p>
          <select v-model="data.setting.form.showBackgroundImage" style="width: 200px;margin-top: 10px;">
            <option :value="true">显示</option>
            <option :value="false">不显示</option>
          </select>
          <p style="font-size: 14px;margin-top: 10px">GPT Key</p>
          <input v-model="data.setting.form.gptKey" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />
        </div>
        <div class="form-footer">
          <button v-if="data.appStarted" style="margin-right: 10px;" @click="controlForm('setting',0)">取消</button>
          <button @click="saveSetting">保存</button>
        </div>
      </div>
    </div>

    <div id="new" class="form-bg" style="display: none;">
      <div class="form-window">
        <div class="form-title">
          新建表格文件
          <svg @click="controlForm('new',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content">
          <p style="font-size: 14px;">文件名</p>
          <input v-model="data.new.newFileName" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />
          <p style="font-size: 14px;margin-top: 10px;">保存位置（{{ data.new.currentFolderPath }}）</p>
          <selectFolder @path-changed="pathChanged" :parent-folder-path="data.setting.readData.saveDrive" v-if="data.setting.readData!=undefined && data.new.showSelectFolder" style="margin-top: 5px;width: 100%;height: 350px;"></selectFolder>
        </div>
        <div class="form-footer">
          <button style="float:left;margin-right: 0px;" @click="()=>{
                data.new.newFolderName = ''
                controlForm('newFolder',1)
              }">新建文件夹</button>
          <button style="margin-right: 10px;" @click="controlForm('new',0)">取消</button>
          <button @click="addFile">新建</button>
        </div>
      </div>
    </div>

    <div id="newFolder" class="form-bg" style="display: none;">
      <div class="form-window" style="height: 240px;">
        <div class="form-title">
          新建文件夹
          <svg @click="controlForm('newFolder',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content">
          <p style="font-size: 14px;">文件夹名</p>
          <input v-model="data.new.newFolderName" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />
          <p style="margin-top: 5px;">文件夹将被新建在{{ ' ' + data.new.currentFolderPath + ' ' }}目录下。</p>
        </div>
        <div class="form-footer">
          <button style="margin-right: 10px;" @click="controlForm('newFolder',0)">取消</button>
          <button @click="addFolder">新建</button>
        </div>
      </div>
    </div>

    <div id="move" class="form-bg" style="display: none;">
      <div class="form-window" style="height: 535px;">
        <div class="form-title">
          移动
          <svg @click="controlForm('move',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content">
          <p style="font-size: 14px;margin-top: 0px;">移动到（{{ data.new.currentFolderPath }}）</p>
          <selectFolder @path-changed="pathChanged" :parent-folder-path="data.setting.readData.saveDrive" v-if="data.setting.readData!=undefined && data.new.showSelectFolder" style="margin-top: 5px;width: 100%;height: 350px;"></selectFolder>
        </div>
        <div class="form-footer">
          <button style="margin-right: 10px;" @click="controlForm('move',0)">取消</button>
          <button @click="fileList_moveClick">移动</button>
        </div>
      </div>
    </div>

    <div id="newSheet" class="form-bg" style="display: none;">
      <div class="form-window" style="height: 215px;">
        <div class="form-title">
          新建表格
          <svg @click="controlForm('newSheet',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content">
          <p style="font-size: 14px;margin-top: 0px;">表格名称</p>
          <input v-model="data.new.newFileName" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />
        </div>
        <div class="form-footer">
          <button style="margin-right: 10px;" @click="controlForm('newSheet',0)">取消</button>
          <button @click="sheet_newSheet">新建</button>
        </div>
      </div>
    </div>

    <div id="rename" class="form-bg" style="display: none;">
      <div class="form-window" style="height: 215px;">
        <div class="form-title">
          重命名
          <svg @click="controlForm('rename',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content">
          <p v-if="data.rightMenu.data!=undefined && (data.rightMenu.name=='fileListFile' || data.rightMenu.name=='fileListDir')" style="font-size: 14px;margin-top: 0px;">{{data.rightMenu.data.split('.')[data.rightMenu.data.split('.').length-1]=='JSON'?'新文件名':'新文件夹名'}}</p>
          <input v-model="data.new.newFileName" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />
        </div>
        <div class="form-footer">
          <button style="margin-right: 10px;" @click="controlForm('rename',0)">取消</button>
          <button @click="fileList_rename">重命名</button>

        </div>
      </div>
    </div>

    <div id="addCol" class="form-bg" style="display: none;">
      <div class="form-window" style="height: auto;max-height: 80vh">
        <div class="form-title">
          新建列
          <svg @click="controlForm('addCol',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content">
          <p style="font-size: 14px;margin-top: 0px;">标题</p>
          <input v-model="data.addCol.name" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />
          <p style="font-size: 14px;margin-top: 10px;">类型</p>
          <select @change="()=>{
            data.addCol.sumType = 'none'
          }" v-model="data.addCol.type" style="min-width: 200px;margin-top: 10px;">
            <option value="text">文本</option>
            <option value="number">数字</option>
            <option value="date">日期</option>
            <option value="sum">行求和</option>
            <option value="sumdiv">行求平均数</option>
          </select>
          <div v-if="data.addCol.type=='sum' || data.addCol.type=='sumdiv'">
            <p style="font-size: 14px;margin-top: 10px;">{{ data.addCol.type=='sum'?'求和':'求平均数' }}的列</p>
            <div style="width: 100%;border-radius: 5px;background-color: rgb(242,242,242);padding: 10px 10px;overflow-y: auto;height: auto;max-height: 300px;margin-top: 10px">
              <p v-if="data.addCol.selectCol.length==0">当前表格没有可求和的列。</p>
              <table v-if="data.addCol.selectCol.length!=0">
                <tr v-for="(item) in data.addCol.selectCol">
                  <td><input style="display: inline-block;position: relative;left: 16px;top: 1.5px" v-model="item.selected" :title="item.name" type="checkbox" /></td>
                  <td style="text-align: left">{{'第' + item.index + '列-' + item.name}}</td>
                </tr>
              </table>
            </div>
          </div>
          <div>
            <p style="font-size: 14px;margin-top: 10px;">底部合计</p>
            <select v-model="data.addCol.sumType" style="min-width: 200px;margin-top: 10px;">
              <option value="none">不显示</option>
              <option v-if="data.addCol.type!='text' && data.addCol.type!='date'" value="sum">列求和</option>
              <option v-if="data.addCol.type!='text' && data.addCol.type!='date'" value="sumdiv">列求平均数</option>
              <option v-if="data.addCol.type=='number' || data.addCol.type=='text'" value="count">列统计非空单元格个数</option>
            </select>
          </div>
        </div>
        <div class="form-footer">
          <button style="margin-right: 10px;" @click="controlForm('addCol',0)">取消</button>
          <button @click="()=>{
            if(data.addCol.name==''){
              setWarningFormTitleAndContentAndShowForm('提示','请输入标题。',false,()=>{})
              return
            }
            let flag = false
            for(let i=0;i<data.addCol.selectCol.length;i++){
              if(data.addCol.selectCol[i].selected==true){
                flag = true
                break
              }
            }
            if(data.addCol.type=='text' || data.addCol.type=='number' || data.addCol.type=='date'){
              flag = true
            }
            if(flag==false){
              setWarningFormTitleAndContentAndShowForm('提示','请选择至少一个参与求和或求平均数的列。',false,()=>{})
              return
            }
            operationStackAppend()
            let key = getColumnsNewKey()
            let sumCols = []
            for(let i=0;i<data.addCol.selectCol.length;i++){
              if(data.addCol.selectCol[i].selected==false)continue
              if(data.addCol.selectCol[i].keyName==key)continue
              sumCols.push(data.addCol.selectCol[i].keyName)
            }
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[key] = {
              name:data.addCol.name,
              type:data.addCol.type,
              sumType:data.addCol.sumType,
              sumCols:sumCols
            }
            for(let i=0;i<data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length;i++){
              let defaultValue = undefined
              if(data.addCol.type=='text'){
                defaultValue = ''
              }else{
                defaultValue = 0
              }
              data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i][key] = defaultValue
            }
            setCurrentFileUnsave()
            controlForm('addCol',0)
          }">新建</button>

        </div>
      </div>
    </div>

    <div id="input" class="form-bg" style="display: none;">
      <div class="form-window" style="height: 215px;">
        <div class="form-title">
          {{ data.input.title }}
          <svg @click="controlForm('input',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content">
          <p style="font-size: 14px;margin-top: 0px;">{{ data.input.subTitle }}</p>
          <input v-model="data.input.text" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />
        </div>
        <div class="form-footer">
          <button style="margin-right: 10px;" @click="controlForm('input',0)">取消</button>
          <button @click="data.input.callBack()">确定</button>

        </div>
      </div>
    </div>

    <div id="addRow" class="form-bg" style="display: none;">
      <div class="form-window" style="height: auto">
        <div class="form-title">
          {{ data.addRow.isEdit?'编辑记录':'插入' }}
          <svg @click="controlForm('addRow',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content" style="height: auto;max-height: calc(80vh - 56.4px - 72.4px);overflow-y: auto">
          <div v-for="(item,index) in data.addRow.inputValues" style="margin-bottom: 10px">
            <p style="font-size: 14px;">{{ '第' + (index+1).toString() + '列-' + item.name }}</p>
<!--            <input v-if="item.type=='text'" v-model="item.value" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />-->
            <textarea v-if="item.type=='text'" v-model="item.value" style="min-height: 128px;max-height: 128px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;min-width: 100%;max-width: 100%;font-family: 微软雅黑"></textarea>
            <input v-if="item.type=='number'" v-model="item.value" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />
            <input v-if="item.type=='date'" v-model="item.value" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;font-family: 微软雅黑" type="date" />
          </div>
        </div>
        <div class="form-footer">
          <button style="margin-right: 10px;" @click="controlForm('addRow',0)">取消</button>
          <button @click="()=>{
            if(data.addRow.isEdit){
              operationStackAppend()
              if(rowCheckInput(data.addRow.inputValues)==false){
                return
              }
              for(let i=0;i<data.addRow.inputValues.length;i++){
                data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[data.rightMenu.data.dataIndex][data.addRow.inputValues[i].keyName] = data.addRow.inputValues[i].value
              }
              data.addRow.isEdit = false
              controlForm('addRow',0)
              setCurrentFileUnsave()
              return
            }
            if(rowCheckInput(data.addRow.inputValues)==false){
              return
            }
            let re = {}
            for(let i=0;i<data.addRow.inputValues.length;i++){
              re[data.addRow.inputValues[i].keyName] = data.addRow.inputValues[i].value
            }
            for(let i=0;i<this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].data.length;i++){
              if(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].data[i]['KpAF0']>=re['KpAF0']){
                this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].data[i]['KpAF0']+=1
              }
            }
            operationStackAppend()
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data = tool.arrInsert(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data,re['KpAF0']-1,re)
            setCurrentFileUnsave()
            controlForm('addRow',0)
          }">{{ data.addRow.isEdit?'保存':'插入' }}</button>
        </div>
      </div>
    </div>

    <div id="editCol" class="form-bg" style="display: none;">
      <div class="form-window" style="height: auto">
        <div class="form-title">
          编辑列
          <svg @click="controlForm('editCol',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content" style="height: auto;max-height: calc(80vh - 56.4px - 72.4px);overflow-y: auto">
          <p v-if="data.editCol.isIndexCol==false" style="font-size: 14px;">标题</p>
          <input v-if="data.editCol.isIndexCol==false" v-model="data.editCol.name" style="height: 32px;width: 100%;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" />
          <p style="font-size: 14px;margin-top: 10px;">底部合计</p>
          <select v-model="data.editCol.sumType" style="min-width: 200px;margin-top: 10px;">
            <option value="none">不显示</option>
            <option v-if="data.editCol.isIndexCol" value="num">显示记录数</option>
            <option v-if="data.editCol.type!='text' && data.editCol.type!='date' && data.editCol.isIndexCol==false" value="sum">列求和</option>
            <option v-if="data.editCol.type!='text' && data.editCol.type!='date' && data.editCol.isIndexCol==false" value="sumdiv">列求平均数</option>
            <option v-if="(data.editCol.type=='number' || data.editCol.type=='text') && data.editCol.isIndexCol==false" value="count">列统计非空单元格个数</option>
          </select>
          <div v-if="data.editCol.isIndexCol==false">
            <p style="font-size: 14px;margin-top: 10px;margin-bottom: 0px">切换列类型（当前列类型：{{ data.editCol.type=='text'?('文本'):(data.editCol.type=='number'?('数字'):(data.editCol.type=='date'?('日期'):(data.editCol.type=='sum'?('行求和'):('行求平均数')))) }}）</p>
            <button v-if="data.editCol.type=='number'" @click="()=>{
              setWarningFormTitleAndContentAndShowForm('提示','此列类型切换为文本之后，列底部合计不能显示求和和求平均数，是否立刻切换？',true,()=>{
                operationStackAppend()
                data.editCol.type = 'text'
                data.editCol.sumType = 'none'
                for(let i=0;i<data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length;i++){
                  data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i][data.editCol.keyName] = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i][data.editCol.keyName].toString()
                }
                data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.editCol.keyName].type = 'text'
                setCurrentFileUnsave()
              })
            }">切换为文本类型</button>
            <button v-if="data.editCol.type=='text'" @click="()=>{
              setWarningFormTitleAndContentAndShowForm('提示','此列类型切换为数字之后，非数字的文本将被替换为数字0，是否立刻切换？',true,()=>{

                operationStackAppend()
                data.editCol.type = 'number'
                data.editCol.sumType = 'none'
                for(let i=0;i<data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data.length;i++){
                    try{
                      let re = JSON.parse(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i][data.editCol.keyName].toString())
                      if(typeof(re)!='number'){
                        data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i][data.editCol.keyName] = 0
                      }else{
                        data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i][data.editCol.keyName] = re
                      }
                    }catch(e){
                      data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[i][data.editCol.keyName] = 0
                    }
                }
                data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.editCol.keyName].type = 'number'
                setCurrentFileUnsave()
              })
            }">切换为数字类型</button>
          </div>
          <p style="font-size: 14px;margin-top: 10px">最小列宽（{{ (data.editCol.minWidth=='' || data.editCol.minWidth==undefined || data.editCol.minWidth==0)?'自动调整':(data.editCol.minWidth.toString() + 'px') }}）</p>
          <input v-model="data.editCol.minWidth" :style="{
            width:(data.editCol.minWidth=='0' || data.editCol.minWidth=='' || data.editCol.minWidth==undefined)?'100%':'calc(100% - 132px)'
          }" placeholder="请输入最小列宽，不输入内容、输入0均表示自动调整列宽" style="height: 32px;padding: 10px 10px;border-radius: 3px;background-color: rgb(242,242,242);border: 0;margin-top: 5px;" type="text" /><button @click="data.editCol.minWidth=''" v-if="!(data.editCol.minWidth=='0' || data.editCol.minWidth=='' || data.editCol.minWidth==undefined)" style="margin: 0 0;margin-left: 10px;position: relative;top: 1px">自动调整列宽</button>
          <div v-if="data.editCol.minWidth!='' && tool.getNumberFromString(data.editCol.minWidth)!=-1 && tool.getNumberFromString(data.editCol.minWidth)!=0" style="margin-top: 10px;border: 1px solid rgb(230,230,230);border-radius: 3px;font-size: 14px;text-align: center;padding-top: 5px;padding-bottom: 5px" :style="{
            width:tool.getNumberFromString(data.editCol.minWidth)==0?('250px'):(tool.getNumberFromString(data.editCol.minWidth).toString() + 'px')
          }">
            {{ tool.getNumberFromString(data.editCol.minWidth)==0?('根据实际情况自动调整'):('此方块宽度为' + tool.getNumberFromString(data.editCol.minWidth).toString() + 'px') }}
          </div>
          <p v-if="data.editCol.minWidth!='' && tool.getNumberFromString(data.editCol.minWidth)==-1" style="font-style: 14px;color: red">最小列宽应为正整数。</p>
        </div>
        <div class="form-footer">
          <button style="margin-right: 10px;" @click="controlForm('editCol',0)">取消</button>
          <button @click="()=>{
            operationStackAppend()
            if(data.editCol.name==''){
              setWarningFormTitleAndContentAndShowForm('提示','请输入标题。',false,()=>{})
              return
            }
            if(data.editCol.minWidth!=''){
              try{
                data.editCol.minWidth = JSON.parse(data.editCol.minWidth)
                if(data.editCol.minWidth<0){
                  throw new Error('')
                }
                if(parseInt(data.editCol.minWidth)!=data.editCol.minWidth){
                  throw new Error('')
                }
              }catch (e) {
                setWarningFormTitleAndContentAndShowForm('提示','最小列宽应为正整数。',false,()=>{})
                return
              }
            }
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].name = data.editCol.name
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].sumType = data.editCol.sumType
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].type = data.editCol.type
            data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[data.rightMenu.data.columnKey].minWidth = data.editCol.minWidth
            setCurrentFileUnsave()
            controlForm('editCol',0)
            refreshTopAndBottom()
          }">保存</button>
        </div>
      </div>
    </div>

    <div id="sort" class="form-bg" style="display: none;">
      <div class="form-window" style="height: auto">
        <div class="form-title">
          排序
          <svg @click="controlForm('sort',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content" style="height: auto;max-height: calc(80vh - 56.4px - 72.4px);overflow-y: auto">
          <p style="font-size: 14px;margin-top: 10px;">索引列</p>
          <select v-model="data.sort.selectColumn" style="min-width: 200px;margin-top: 10px;">
            <option v-for="item in data.sort.columns" :value="item.key">{{ item.name }}</option>
          </select>
          <p style="font-size: 14px;margin-top: 10px;">顺序</p>
          <select v-model="data.sort.isRe" style="min-width: 200px;margin-top: 10px;">
            <option :value="false">升序</option>
            <option :value="true">降序</option>
          </select>
        </div>
        <div class="form-footer">
          <button style="margin-right: 10px;" @click="controlForm('sort',0)">取消</button>
          <button @click="async ()=>{
            operationStackAppend()
            showLoading()
            let cks = Object.keys(currentFile.data.sheets[currentFile.system.currentSheetIndex-1].columns)
            let cki = 0
            let isNum = false
            for(let i=0;i<cks.length;i++){
              if(cks[i]==data.sort.selectColumn){
                cki = i
                if(currentFile.data.sheets[currentFile.system.currentSheetIndex-1].columns[data.sort.selectColumn].type=='number'){
                  isNum = true
                }
                break
              }
            }
            let tablet = document.getElementById('mainTable').getElementsByTagName('tr')
            let indexArr = []
            for(let i=1;i<tablet.length;i++){
              if(isNum){
                indexArr.push(JSON.parse(tablet[i].getElementsByTagName('td')[cki].innerText))
              }else{
                indexArr.push(tablet[i].getElementsByTagName('td')[cki].innerText)
              }
            }
            let ct = currentFile.data.sheets[currentFile.system.currentSheetIndex-1].data
            ct = tool.sortArray(indexArr,ct,isNum)
            let ct2 = []
            if(data.sort.isRe){
              for(let i=ct.length-1;i>=0;i--){
                ct2.push(ct[i])
              }
            }else{
              ct2 = ct
            }
            await tool.waitSeconds(1)
            currentFile.data.sheets[currentFile.system.currentSheetIndex-1].data = ct2
            for(let i=0;i<currentFile.data.sheets[currentFile.system.currentSheetIndex-1].data.length;i++){
              currentFile.data.sheets[currentFile.system.currentSheetIndex-1].data[i]['KpAF0'] = i+1
            }
            hideLoading()
            setCurrentFileUnsave()
            controlForm('sort',false)
          }">立刻排序</button>
        </div>
      </div>
    </div>

    <div id="warning" class="form-bg" style="display: none;">
      <div class="form-window" style="height: auto;">
        <div class="form-title">
          <span id="warningFormTitle">{{ data.warning.title }}</span>
          <svg v-if="data.warning.showCancel" @click="controlForm('warning',false)" style="position: absolute;right: 20px;top: 20px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
        </div>
        <div class="form-content">
          <p style="font-size: 14px;">{{ data.warning.content }}</p>
        </div>
        <div class="form-footer">
          <button v-if="data.warning.showCancel" @click="controlForm('warning',0)" style="margin-right: 10px">{{data.warning.buttonTexts.cancel}}</button>
          <button v-if="data.warning.showSecondary" @click="controlForm('warning',0);data.warning.secondaryCallBack()" style="margin-right: 10px">{{data.warning.buttonTexts.secondary}}</button>
          <button @click="controlForm('warning',0);data.warning.callBack()">{{data.warning.buttonTexts.ok}}</button>
        </div>
      </div>
    </div>

    <div id="loading" class="form-bg" style="display: none;">
      <div class="form-window" style="width: 400px;height: 400px;padding-left: 200px;padding-top: 200px;">
        <div style="text-align: center;transform: translate(-50%,-50%);">
          <div style="--size: 44px; --dot-size: 5px; --dot-count: 6; --color: black; --speed: 1s; --spread: 60deg;display: inline-block;" class="dots">
            <div style="--i: 0;" class="dot"></div>
            <div style="--i: 1;" class="dot"></div>
            <div style="--i: 2;" class="dot"></div>
            <div style="--i: 3;" class="dot"></div>
            <div style="--i: 4;" class="dot"></div>
            <div style="--i: 5;" class="dot"></div>
          </div>
          <p style="font-weight: bold;margin-top: 10px;">请稍后</p>
        </div>
        
      </div>
    </div>

    <div class="white-background" style="width: 100vw;height: calc(100vh - 48px);position: fixed;z-index: 100000;left: 0;bottom: 0;transition-property: opacity,transform,backgroundImage;transition-duration: .2s;
    transition-timing-function: cubic-bezier(0.23, 1, 0.320, 1);background-position: center;background-size: cover;background-color: white"
      :style="{
        opacity:data.showStartPage==2?1:0,
        transform:data.showStartPage==2?'translateX(0%)':'translateX(-256px)',
        display:data.showStartPage==0?'none':'unset',
        backgroundImage:data.backgroundImage
      }"

    >


      <div style="width: 100%;height: 100%;background-color: rgba(255,255,255,0.5);">
        <div style="padding: 40px 44px;">
          <p style="font-size: 28px;font-weight: bold;">{{ getGreetingBasedOnTime() }}</p>
          <div style="margin-top: 40px;display: inline-block;float: left;">
            <p style="font-weight: bold;display: inline-block;">操作</p><br>

            <div style="display: inline-block;">
              <button v-if="data.files.length!=0" style="width: 100px;margin-right: 10px;backdrop-filter: blur(10px);background-color: rgba(255,255,255,0.3);" @click="hideStartPage">
                <svg style="margin-top: 15px;margin-bottom: 15px;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M7 4a1 1 0 011 1v38a1 1 0 01-1 1H5a1 1 0 01-1-1V5a1 1 0 011-1h2zm15.669 11.114a1.4 1.4 0 01.331.904V22h20a1 1 0 011 1v2a1 1 0 01-1 1H23v5.982a1.4 1.4 0 01-2.304 1.068l-9.433-7.981a1.4 1.4 0 010-2.138l9.433-7.981a1.4 1.4 0 011.973.164z" fill="currentColor"/></svg>
                <p>返回</p>
              </button>

              <button style="width: 100px;margin-right: 10px;backdrop-filter: blur(10px);background-color: rgba(255,255,255,0.3);" @click="async ()=>{
            newButtonClick()
          }">
                <svg style="margin-top: 15px;margin-bottom: 15px;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M24.718 4c.446 0 .608.046.77.134.163.087.291.215.378.378.088.163.134.324.134.77V22h16.718c.446 0 .607.046.77.134.163.087.291.215.378.378.088.162.134.324.134.77v1.436c0 .446-.046.608-.134.77a.908.908 0 01-.378.378c-.163.088-.324.134-.77.134H26L26 42.718c0 .446-.046.607-.134.77a.908.908 0 01-.378.378c-.162.088-.324.134-.77.134h-1.436c-.446 0-.608-.046-.77-.134a.908.908 0 01-.378-.378c-.088-.163-.134-.324-.134-.77L21.999 26 5.282 26c-.446 0-.607-.046-.77-.134a.908.908 0 01-.378-.378c-.088-.162-.134-.324-.134-.77v-1.436c0-.446.046-.608.134-.77a.908.908 0 01.378-.378c.163-.088.324-.134.77-.134L22 21.999V5.282c0-.446.046-.607.134-.77a.908.908 0 01.378-.378c.162-.088.324-.134.77-.134h1.436z" fill="currentColor"/></svg>
                <p>新建</p>
              </button>

              <button style="width: 100px;margin-right: 10px;backdrop-filter: blur(10px);background-color: rgba(255,255,255,0.3);" @click="controlForm('setting',true)">
                <svg style="margin-top: 15px;margin-bottom: 15px;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M28.767 2.71A21.685 21.685 0 0024 2.182c-1.618 0-3.215.178-4.768.528l-.897.202-2.795 6.373-6.883-.752-.622.677a21.82 21.82 0 00-4.772 8.295l-.27.872L7.089 24 3.53 28.886l-.537.737.27.872a21.82 21.82 0 004.772 8.295l.622.677 6.883-.752 2.426 5.531.369.842.896.202c1.554.35 3.15.528 4.769.528 1.618 0 3.214-.178 4.767-.528l.897-.201 2.796-6.374 6.883.752.622-.677a21.82 21.82 0 004.771-8.294l.27-.871L40.912 24l3.559-4.887.537-.738-.27-.87a21.82 21.82 0 00-4.772-8.295l-.622-.677-6.883.752-2.796-6.373-.897-.202zM20.916 6.08A18.109 18.109 0 0124 5.818c1.043 0 2.073.089 3.084.263l2.24 5.109a2.976 2.976 0 003.05 1.763l5.507-.603a18.182 18.182 0 013.096 5.382l-3.289 4.516a2.976 2.976 0 000 3.504l3.289 4.516a18.182 18.182 0 01-3.096 5.382l-5.508-.603a2.976 2.976 0 00-3.049 1.763l-2.24 5.109c-1.011.174-2.041.263-3.084.263-1.043 0-2.073-.089-3.084-.263l-2.24-5.108a2.976 2.976 0 00-3.05-1.764l-5.508.603a18.185 18.185 0 01-3.095-5.383l3.289-4.515c.76-1.044.76-2.46 0-3.504l-3.29-4.515a18.186 18.186 0 013.096-5.383l5.509.603a2.976 2.976 0 003.049-1.763l2.24-5.109zM24 14.91c5.004 0 9.056 4.072 9.056 9.091 0 5.019-4.052 9.09-9.056 9.09S14.944 29.02 14.944 24s4.052-9.09 9.056-9.09zM18.58 24c0-3.014 2.429-5.454 5.42-5.454s5.42 2.44 5.42 5.454-2.429 5.455-5.42 5.455-5.42-2.44-5.42-5.455z" fill="currentColor"/></svg>
                <p>设置</p>
              </button>

              <button style="width: 100px;margin-right: 10px;backdrop-filter: blur(10px);background-color: rgba(255,255,255,0.3);" @click="importFile">
                <svg style="margin-top: 15px;margin-bottom: 15px;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M31 4a1 1 0 011 1v2a1 1 0 01-1 1H8v32h23a1 1 0 011 1v2a1 1 0 01-1 1H6a2 2 0 01-2-2V6a2 2 0 012-2h25zM17.513 22.436l7.778-7.778a1 1 0 011.32-.084l.095.084 1.414 1.414a1 1 0 01.083 1.32l-.083.094-4.602 4.601 18.573-.082c.52.002.942.39.998.888l.006.116-.008 1.894a1 1 0 01-.88.989l-.115.007-18.207.08 4.235 4.235a1 1 0 01.083 1.32l-.083.094-1.414 1.414a1 1 0 01-1.32.084l-.095-.084-7.778-7.778a2 2 0 01-.117-2.701l.117-.127 7.778-7.778-7.778 7.778z" fill="currentColor"/></svg>
                <p>导入</p>
              </button>
            </div>

          </div>
          <div v-if="data.setting.readData!=undefined" style="margin-top: 40px;width: calc(100% - 490px);display: inline-block;position: absolute;left: 490px;">
            <p style="font-weight: bold;display: inline-block;padding-left:10px">{{ '表格文件（位于' + data.setting.readData.saveDrive.slice(0,data.setting.readData.saveDrive.length-1) + '盘）' }}
              <div title="新建文件夹" @click="newFolderClick" class="button" style="width:25px;height:25px;display:inline-block;text-align:center;padding-top:2.5px;border-radius:3px;position:absolute;right:40px;top:0px">
                <svg width="14" height="14" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M24.718 4c.446 0 .608.046.77.134.163.087.291.215.378.378.088.163.134.324.134.77V22h16.718c.446 0 .607.046.77.134.163.087.291.215.378.378.088.162.134.324.134.77v1.436c0 .446-.046.608-.134.77a.908.908 0 01-.378.378c-.163.088-.324.134-.77.134H26L26 42.718c0 .446-.046.607-.134.77a.908.908 0 01-.378.378c-.162.088-.324.134-.77.134h-1.436c-.446 0-.608-.046-.77-.134a.908.908 0 01-.378-.378c-.088-.163-.134-.324-.134-.77L21.999 26 5.282 26c-.446 0-.607-.046-.77-.134a.908.908 0 01-.378-.378c-.088-.162-.134-.324-.134-.77v-1.436c0-.446.046-.608.134-.77a.908.908 0 01.378-.378c.163-.088.324-.134.77-.134L22 21.999V5.282c0-.446.046-.607.134-.77a.908.908 0 01.378-.378c.162-.088.324-.134.77-.134h1.436z" fill="currentColor"/></svg>
              </div>
              <div title="刷新" @click="refreshFileList" class="button" style="width:25px;height:25px;display:inline-block;text-align:center;padding-top:2.5px;border-radius:3px;position:absolute;right:75px;top:0px">
                <svg width="14" height="14" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M29 20a1 1 0 01-1-1v-2a1 1 0 011-1h6.49c-2.53-3.627-6.732-6-11.49-6-7.732 0-14 6.268-14 14s6.268 14 14 14c6.342 0 11.699-4.217 13.42-10h4.134C39.734 36.017 32.566 42 24 42c-9.941 0-18-8.059-18-18S14.059 6 24 6c5.655 0 10.701 2.608 14.001 6.687L38 7a1 1 0 011-1h2a1 1 0 011 1v11a2 2 0 01-1.85 1.994L40 20H29z" fill="currentColor"/></svg>
              </div>
            </p><br>
            <p style="font-size:14px;margin-top: 10px;padding-left:10px">当前位置：
              <span class="button" @click="fileListBackToStart">{{ data.setting.readData.saveDrive.slice(0,data.setting.readData.saveDrive.length-1) + '盘' }}</span>
              <span v-for="item in data.fileList.path"><span>></span><span class="button" @click="fileListNavItemClick(item)">{{ item.name }}</span></span>
            </p>
            <div v-if="data.fileList.current.length!=0" style="width: 100%;height: calc(100vh - 220px);margin-top: 5px;display: inline-block;overflow-y: auto;padding-right: 40px;padding-left:10px;padding-top:5px;overflow-x:hidden;padding-bottom: 40px">
              <div style="border-radius:5px;box-shadow:0 1px 3px rgba(0,0,0,0.1);overflow:hidden;border:1px solid rgba(255,255,255,0.5);cursor:pointer;background-color: rgba(255,255,255,0.3);backdrop-filter: blur(20px)">
                <div v-for="item in data.fileList.current" :style="{
              display:((item.children!=undefined || getFileExtension(item.name)=='JSON')?'unset':'none')
            }">
                  <div v-if="item.children!=undefined || getFileExtension(item.name)=='JSON'" style="border-bottom:1px solid rgb(230,230,230)">

                    <div class="button" style="padding: 10px 10px;color: black;"
                         @click="()=>{
                    if(item.children==undefined){
                      file_open(data.setting.readData.saveDrive + '\\ExcelEasyData\\' + getCurrentFileListFolderPath() + '\\' + item.name)
                    }else{
                      fileListItemClick(item.name)
                    }
                  }"
                         @mouseup="showRightMenu((item.children==undefined?'fileListFile':'fileListDir'),data.setting.readData.saveDrive + '\\ExcelEasyData\\' + getCurrentFileListFolderPath() + '\\' + item.name)"
                    >

                      <div style="height:35px;width:35px;padding-left: calc(35px /2);padding-top: calc(35px / 2);display: inline-block;float: left;">
                        <svg v-if="item.children!=undefined" style="transform: translate(-50%,-50%);" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd"  clip-rule="evenodd" d="M2 8a2 2 0 012-2h16a2 2 0 012 2v2h21.905c1.157 0 2.095.895 2.095 2v28c0 1.105-.938 2-2.095 2H4.095C2.938 42 2 41.105 2 40V8zm40 6H6v24h36V14z"  fill="currentColor"/></svg>
                        <svg v-if="item.children==undefined" style="transform: translate(-50%,-50%);" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd"  clip-rule="evenodd" d="M34.171 2a2 2 0 011.415.586l5.828 5.831A2 2 0 0142 9.831V44a2 2 0 01-2 2H8a2 2 0 01-2-2V4a2 2 0 012-2h26.171zM32 6H10v36h28V12.038l-5 .001a1 1 0 01-1-1L31.999  6zM33 16a2 2 0 012 2v17a2 2 0 01-2 2H15a2 2 0 01-2-2V18a2 2 0 012-2h18zM20.514 26H16.6v7.4h3.914V26zm10.885 0h-7.885v7.4H31.4l-.001-7.4zm-10.885-6.401L16.6 19.6V23h3.914v-3.401zM31.4   19.6l-7.886-.001V23h7.885l.001-3.4z" fill="currentColor"/></svg>
                      </div>
                      <div style="display:inline-block;margin-left: 10px;width: calc(100% - 50px)">
                        <p class="fileItemTitle" style="font-size: 14px;display:inline-block">{{ (item.children==undefined?getFileNameWithoutExtension(item.name):item.name) }}</p>
                        <p style="color: gray;font-size: 12px;">{{ (item.children==undefined?'修改于':'创建于') + formatDate(item.time) }}</p>
                      </div>
                    </div>
                  </div>

                </div>
              </div>
            </div>
            <div v-if="data.fileList.current.length==0" style="padding-left:10px;margin-top:10px">
              <p>当前文件夹为空。</p>
            </div>
          </div>
        </div>
      </div>


    </div>


    <div style="width: 100vw;height: calc(100vh - 48px);position: relative;">
      <div style="width: 100%;height: 36px;position: relative;padding-left: 72px;box-shadow: 0 1px 3px rgba(0,0,0,0);padding-right: 36px;z-index: 2000;background-color: white;">

        <div @mouseup="showRightMenu('fileBarMenu',undefined)" style="width: 100%;height: 36px;position: absolute;left: 0;top: 0;z-index: 1001"></div>

        <div @mouseup="()=>{
          data.rightMenu.left = 0
          data.rightMenu.top = 96
          data.rightMenu.allowShow = true
          showRightMenu('fileBarMenu',undefined)
        }" style="width: 36px;height: 36px;display: inline-block;min-width: 0;padding: 0 0;position: absolute;left: 0;top: 0;z-index: 1010" class="fileBar-fileItem"
        >
          <svg style="position: absolute;left: 9px;top: 7px" width="20" height="20" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M27 36v4a1 1 0 01-1 1h-4a1 1 0 01-1-1v-4a1 1 0 011-1h4a1 1 0 011 1zm0-24a1 1 0 01-1 1h-4a1 1 0 01-1-1V8a1 1 0 011-1h4a1 1 0 011 1v4zm0 14a1 1 0 01-1 1h-4a1 1 0 01-1-1v-4a1 1 0 011-1h4a1 1 0 011 1v4z" fill="currentColor"/></svg>
        </div>

        <div style="width: 36px;height: 36px;display: inline-block;min-width: 0;padding: 0 0;position: absolute;right: 0px;top: 0;z-index: 1010" class="fileBar-fileItem" @click="fileBarLeftRightClick(0)"
             :style="{
              display:fileBarLetfRightVisible(false)==false?'none':'unset'
             }"
        >
          <svg style="position: absolute;left: 9px;top: 7px;scale: 0.85;" width="20" height="20" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M32.556 24L16.293 7.737a1 1 0 010-1.415l1.414-1.414a1 1 0 011.414 0L36.8 22.586a2 2 0 010 2.828L19.121 43.092a1 1 0 01-1.414 0l-1.414-1.414a1 1 0 010-1.414L32.556 24z" fill="currentColor"/></svg>
        </div>

        <div style="width: 36px;height: 36px;display: inline-block;min-width: 0;padding: 0 0;position: absolute;left: 36px;top: 0;z-index: 1010" class="fileBar-fileItem"
             :style="{
              display:fileBarLetfRightVisible(true)==false?'none':'unset'
             }" @click="fileBarLeftRightClick(1)">
          <svg style="position: absolute;left: 9px;top: 7px;scale: 0.85" width="20" height="20" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M15.272 24l16.263 16.264a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L11.03 25.414a2 2 0 010-2.828L28.706 4.908a1 1 0 011.414 0l1.415 1.414a1 1 0 010 1.415L15.271 24z" fill="currentColor"/></svg>
        </div>

        <div id="fileBarWrapper" style="height: 36px;width: 100%;">

          <div style="width: 10000px;position: relative;z-index: 1002">
            <div id="fileBar" style="display: inline-block;position: absolute;top: 0;"
                 :style="{
                left:data.fileBar.left.toString() + 'px',
                transition:data.fileBar.showAnimation?'left .5s cubic-bezier(0.23, 1, 0.320, 1)':'unset'
              }"
            >

              <div :class="data.currentFile.path==item.system.path?['fileBar-fileItem fileBar-fileItem-selected']:['fileBar-fileItem']" v-for="item in data.files" class="fileBar-fileItem" style="display: inline-block;position: relative">
                <div style="position: absolute;width: 100%;height: 100%;left: 0;top: 0" @click="file_fileBarItemClick(item.system.path)" :title="fileBarTitleGetPath(item.system.path)"
                     @mouseup="()=>{
              showRightMenu('fileBarFileItemMenu',item.system.path)
            }"
                >

                </div>
                {{ (item.system.unsave?'*':'') + item.data.name }}
                <svg @click="data.rightMenu.data = item.system.path;file_closeSingle()" style="position: absolute;right: 10px;top: 11px;cursor: pointer" width="14" height="14" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
              </div>

            </div>
          </div>



        </div>

      </div>
      <div v-if="data.currentFile.index!=0 && data.files[data.currentFile.index-1]!=undefined" style="width: 100%;height: calc(100vh - 48px - 36px - 0px);position: relative;">

        <div style="width: calc(100%);height: 36px;text-align: center;top: 0px;box-shadow: 0 0px 3px rgba(0,0,0,0.2);z-index: 1000;border-top: 1px solid rgb(230,230,230);border-bottom: 1px solid rgb(230,230,230);position: relative;background-color: #fff;">

          <div v-if="data.files[data.currentFile.index-1].system.search.show==false">

            <div class="opBarItem" @click="()=>{
                this.data.rightMenu.data = data.files[data.currentFile.index-1].system.path
                file_saveSingle()
              }"
                 :class="data.files[data.currentFile.index-1].system.unsave==false?['opBarItem-disabled']:['opBarItem']"
            >
              <svg width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M32.38 4a4 4 0 012.832 1.175l7.62 7.64A4 4 0 0144 15.64V42a2 2 0 01-2 2H6a2 2 0 01-2-2V6a2 2 0 012-2h26.38zM12 8H8v32h32V16.83a2 2 0 00-.584-1.412l-6.81-6.83A2 2 0 0031.19 8H31v14a2 2 0 01-2 2H14a2 2 0 01-2-2V8zm15 0H16v12h3.999L20 14a1 1 0 011-1h2a1 1 0 011 1v6h3V8z" fill="currentColor"/></svg>
              <span>保存</span>

            </div>

            <div class="opBarItem" @click="async ()=>{
                if(data.files[data.currentFile.index-1].system.unsave==true){
                  setWarningFormTitleAndContentAndShowForm('准备导出','当前文件尚未保存，是否保存并导出当前文件？',true,async ()=>{
                    data.rightMenu.data = data.files[data.currentFile.index-1].system.path
                    await file_saveSingle()
                    await exportExcelFileFromPath(data.files[data.currentFile.index-1].system.path)
                  })
                  return
                }
                exportExcelFileFromPath(data.files[data.currentFile.index-1].system.path)
              }"
                 :class="['opBarItem']"
            >
              <svg width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M31 4a1 1 0 011 1v2a1 1 0 01-1 1H8v32h23a1 1 0 011 1v2a1 1 0 01-1 1H6a2 2 0 01-2-2V6a2 2 0 012-2h25zm4.846 10.658l7.778 7.778a2 2 0 010 2.828l-7.778 7.778a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l4.235-4.235-18.206-.08a1 1 0 01-.996-.996l-.008-1.894a1 1 0 01.88-.997l.125-.007 18.572.082-4.602-4.601a1 1 0 010-1.414l1.414-1.414a1 1 0 011.415 0z" fill="currentColor"/></svg>
              <span>导出</span>

            </div>

            <div class="opBarItem" @click="()=>{
                withDrawClick()

              }"
                 :class="data.files[data.currentFile.index-1].system.operationStackIndex==0?['opBarItem-disabled']:['opBarItem']"
            >
              <svg width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23 15.37c12.15 0 22 9.992 22 22.318 0 .766-.083 1.795-.248 3.088a.5.5 0 01-.92.202l-.317-.496c-.304-.467-.57-.847-.798-1.14-3.586-4.616-9.448-7.225-17.17-7.331L23 32.007v8.975c0 .562-.46 1.018-1.029 1.018-.273 0-.534-.107-.727-.298L4.452 25.08a1.516 1.516 0 010-2.16L21.244 6.298a1.036 1.036 0 011.455 0c.193.191.301.45.301.72v8.352zm0 4h-4v-5.222L9.047 24 19 33.852v-5.845h6.052c5.805 0 10.882 1.263 14.972 3.713C37.59 24.529 30.88 19.37 23 19.37z" fill="currentColor"/></svg>
              <span>撤销</span>

            </div>

            <div class="opBarItem" @click="()=>{
                reDo()
              }"
                 :class="data.files[data.currentFile.index-1].system.operationStackIndex>=(data.files[data.currentFile.index-1].system.operationStack.length-1)?['opBarItem-disabled']:['opBarItem']"
            >
              <svg style="transform: rotateY(180deg)" width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23 15.37c12.15 0 22 9.992 22 22.318 0 .766-.083 1.795-.248 3.088a.5.5 0 01-.92.202l-.317-.496c-.304-.467-.57-.847-.798-1.14-3.586-4.616-9.448-7.225-17.17-7.331L23 32.007v8.975c0 .562-.46 1.018-1.029 1.018-.273 0-.534-.107-.727-.298L4.452 25.08a1.516 1.516 0 010-2.16L21.244 6.298a1.036 1.036 0 011.455 0c.193.191.301.45.301.72v8.352zm0 4h-4v-5.222L9.047 24 19 33.852v-5.845h6.052c5.805 0 10.882 1.263 14.972 3.713C37.59 24.529 30.88 19.37 23 19.37z" fill="currentColor"/></svg>
              <span>重做</span>

            </div>

            <div class="opBarItem" @click="()=>{
                if(Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns).length==1){
                  setWarningFormTitleAndContentAndShowForm('提示','当前表格没有可以输入内容的列，所以不能插入。',false,()=>{})
                  return
                }
                data.addRow.inputValues = addRow_getCurrentSheetCols()
                data.addRow.isEdit = false
                controlForm('addRow',1)
              }">
              <svg width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M45 5a1 1 0 00-1-1H4a1 1 0 00-1 1v2a1 1 0 001 1h27.999v23H5a2 2 0 00-2 2v9a2 2 0 002 2h39a1 1 0 001-1V5zM7 35h33.999L41 40H7v-5zM41 8h-5l-.002 23h5.001L41 8zm-22.4 6.2a1 1 0 00-1.6.8v2H4a1 1 0 00-1 1v2a1 1 0 001 1h12.999L17 23a1 1 0 001.6.8l5.333-4a1 1 0 000-1.6l-5.333-4z" fill="currentColor"/></svg>
              <span>插入记录</span>

            </div>

            <div class="opBarItem" @click="()=>{
                addCol_getCurrentSheetCols()
                data.addCol.name = ''
                data.addCol.type = 'text'
                data.addCol.sumType = 'none'
                controlForm('addCol',1)
              }">
              <svg width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M24.718 4c.446 0 .608.046.77.134.163.087.291.215.378.378.088.163.134.324.134.77V22h16.718c.446 0 .607.046.77.134.163.087.291.215.378.378.088.162.134.324.134.77v1.436c0 .446-.046.608-.134.77a.908.908 0 01-.378.378c-.163.088-.324.134-.77.134H26L26 42.718c0 .446-.046.607-.134.77a.908.908 0 01-.378.378c-.162.088-.324.134-.77.134h-1.436c-.446 0-.608-.046-.77-.134a.908.908 0 01-.378-.378c-.088-.163-.134-.324-.134-.77L21.999 26 5.282 26c-.446 0-.607-.046-.77-.134a.908.908 0 01-.378-.378c-.088-.162-.134-.324-.134-.77v-1.436c0-.446.046-.608.134-.77a.908.908 0 01.378-.378c.163-.088.324-.134.77-.134L22 21.999V5.282c0-.446.046-.607.134-.77a.908.908 0 01.378-.378c.162-.088.324-.134.77-.134h1.436z" fill="currentColor"/></svg>
              <span>新建列</span>
            </div>

            <div class="opBarItem" @click="()=>{
                data.files[data.currentFile.index-1].system.search.show = true
                let re = Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)
                let re0 = []
                for(let i=0;i<re.length;i++){
                  re0.push({
                    key:re[i],
                    name:data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[re[i]].name
                  })
                }
                data.files[data.currentFile.index-1].system.search.selectKeys = re0
                data.files[data.currentFile.index-1].system.search.key = re0[0].key
                data.files[data.currentFile.index-1].system.search.showResult = false
              }">
              <svg width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M22 5c9.389 0 17 7.611 17 17 0 3.549-1.087 6.844-2.947 9.57l6.782 6.782a1 1 0 010 1.414l-1.697 1.698a1 1 0 01-1.414 0l-6.604-6.605A16.934 16.934 0 0122 39c-9.389 0-17-7.611-17-17S12.611 5 22 5zm0 4.2C14.93 9.2 9.2 14.93 9.2 22S14.93 34.8 22 34.8 34.8 29.07 34.8 22 29.07 9.2 22 9.2z" fill="currentColor"/></svg>
              <span>搜索</span>
            </div>

            <div class="opBarItem" @click="()=>{
                let cfks = Object.keys(currentFile.data.sheets[currentFile.system.currentSheetIndex-1].columns)
                let re = []
                for(let i=0;i<cfks.length;i++){
                  re.push({
                    key:cfks[i],
                    name: '第' + (i+1).toString() + '列-' + currentFile.data.sheets[currentFile.system.currentSheetIndex-1].columns[cfks[i]].name
                  })
                }
                data.sort.columns = re
                data.sort.selectColumn = re[0].key
                data.sort.isRe = false
                controlForm('sort',true)
              }">
              <svg width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path d="M14.38 28.456a.278.278 0 01.213-.456h18.814c.235 0 .364.275.213.456l-9.108 10.93a.667.667 0 01-1.024 0l-9.108-10.93zM14.38 19.544a.278.278 0 00.213.456h18.814a.278.278 0 00.213-.456l-9.108-10.93a.667.667 0 00-1.024 0l-9.108 10.93z" fill="currentColor"/></svg>
              <span>排序</span>
            </div>

            <div class="opBarItem"
                 :style="{
                    background:data.aiAides.show?'linear-gradient(135deg,#b09afd,#3470e8)':'',
                    color:data.aiAides.show?'white':'black',
                 }"
                 @click="async ()=>{
          if(data.aiAides.working)return
                data.aiAides.showAnimation = true
                data.aiAides.show = !data.aiAides.show
                await tool.waitSeconds(0.3)
                data.aiAides.showAnimation = false
                opendedAiAds()
              }">
              <svg width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M31.685 2.026a9.128 9.128 0 00-.993-.02l-.273.012-.283.023-.326.037-.351.053-.28.054-.305.068-.239.06-.41.123-.236.08c-.17.06-.337.125-.503.196l-.3.134-.369.183-.203.11a8.89 8.89 0 00-.663.41l-.266.187-.248.188c-.291.23-.569.477-.83.742l-.052.052c-.145.15-.286.303-.42.462L24 5.344l-.135-.164a9.04 9.04 0 00-.394-.434l-.358-.35a9.035 9.035 0 00-.527-.453l-.27-.206-.266-.187-.228-.15a9.053 9.053 0 00-.437-.26l-.132-.073-.397-.201-.254-.116-.34-.14-.252-.094-.154-.053-.303-.095-.384-.105-.28-.064-.352-.068-.28-.044-.278-.034-.284-.026-.358-.02L17 2a9.1 9.1 0 00-1.023.057l-.218.028-.188.028c-.13.02-.26.044-.39.07l-.145.032-.26.062-.199.053-.312.093-.167.055a8.875 8.875 0 00-1.278.55l-.189.102-.216.124-.328.204-.263.178a8.775 8.775 0 00-.278.205l-.079.06c-.17.133-.335.272-.496.417l-.121.112-.296.289-.224.239-.178.202-.295.364-.201.273-.214.315-.232.378c-.196.337-.37.688-.52 1.05l-.071.175-.149.407c-.07.207-.132.418-.187.633l-.013.049c-.1.4-.174.81-.219 1.23l-.002.03a7.005 7.005 0 00-.605.11l-.248.06-.344.102-.394.14a7 7 0 00-.32.134l-.052.024-.254.122-.198.105-.23.132-.205.128-.244.165-.183.135-.171.134-.334.288-.165.156-.2.204-.125.136-.107.123-.176.215-.172.228-.11.157-.201.312-.103.175-.105.193-.147.295-.107.241-.076.186-.052.14-.082.239-.049.155-.07.251-.029.12c-.1.418-.162.852-.183 1.296L2 17c0 .183.007.365.02.545l.03.29.042.304.072.374c.043.194.094.385.153.573l.11.325.113.29.127.284.079.162c.119.237.251.465.396.685l.09.135.238.326c.26.334.55.645.867.928A10.953 10.953 0 002 29c0 4.355 2.53 8.118 6.201 9.901l.037.163.059.236.094.333.167.493.078.203.13.312.078.17.136.278.099.187.079.143.093.162.168.272.107.163c.077.114.156.226.238.336l.053.071.233.296a9.091 9.091 0 00.514.572l.153.153.277.26.267.229.251.201.17.127.238.17.28.183.368.22.251.138.358.178.317.141.249.102.286.105a9 9 0 00.361.118l.212.061.448.111.298.06.199.035c.13.021.262.04.394.055l.295.03.23.016c.128.008.258.013.39.015L17 46l.308-.005.273-.014.283-.022.326-.037.351-.053.28-.053.308-.07.24-.061.409-.122.233-.079c.17-.06.337-.126.503-.196l.313-.14.208-.101.219-.114.358-.203c.15-.09.297-.183.44-.281l.28-.197.23-.176c.292-.23.57-.477.83-.742l.053-.052c.145-.15.286-.303.42-.462l.134-.164.135.163c.127.15.258.294.394.435l.08.081c.253.256.523.497.805.721l.167.13.293.212.236.158.284.177.221.128.132.072.397.201.24.11.353.146.253.094.233.08.224.068.384.105.274.063.355.069.282.044.28.034.282.026.36.02L31 46l.34-.006.194-.01c.154-.009.307-.021.458-.038l.265-.033.196-.03c.133-.021.266-.046.397-.073l.05-.01.454-.11.34-.1.277-.092.286-.105.25-.102.224-.098.231-.11.348-.18.262-.15.27-.165.308-.206.264-.192.325-.257.244-.21.163-.148.137-.131.16-.16a9.038 9.038 0 00.507-.565l.118-.147.115-.148.174-.239.117-.169.107-.163.2-.326.099-.175.14-.263.094-.19.076-.162.123-.283.13-.331.072-.202.087-.266.061-.208.074-.279.061-.267.267-.134A10.999 10.999 0 0046 29c0-2.557-.873-4.911-2.337-6.779l.2-.186a7 7 0 00.378-.395l.18-.212.145-.182.182-.25.11-.163c.151-.23.29-.47.413-.72l.062-.127.144-.33.097-.245.11-.325c.058-.188.11-.379.152-.573l.052-.262.042-.26.045-.392.02-.326L46 17l-.012-.409a6.991 6.991 0 00-.24-1.46l-.02-.073-.069-.222-.08-.232-.13-.33-.134-.298-.15-.293-.176-.308-.154-.243-.157-.227-.115-.155-.097-.123-.21-.25-.188-.206a7.066 7.066 0 00-.327-.321l-.075-.068-.268-.228-.17-.134a6.838 6.838 0 00-.477-.331l-.092-.058-.115-.069-.32-.177-.284-.14-.22-.098-.277-.11a6.952 6.952 0 00-1.491-.374l-.048-.383-.06-.356-.033-.166-.081-.355-.083-.307-.1-.324-.064-.185-.049-.135-.09-.233-.17-.391-.155-.319-.098-.185a9.113 9.113 0 00-.17-.302l-.086-.14-.21-.326-.11-.159a8.914 8.914 0 00-.213-.288l-.097-.123-.273-.325-.257-.28-.176-.18-.24-.229-.305-.268-.253-.205-.216-.163-.14-.1a9.007 9.007 0 00-.353-.237l-.298-.181a9.027 9.027 0 00-.312-.174l-.032-.016-.395-.196-.185-.083a8.93 8.93 0 00-1.034-.378l-.073-.021-.366-.097-.308-.07a8.953 8.953 0 00-.767-.122l-.367-.035zM8.157 14.12l-.018.004.193-.05.194-.037.121.319a9.009 9.009 0 003.684 4.34 2 2 0 102.078-3.419 4.995 4.995 0 01-2.388-4.743l.02-.18.023-.158.059-.303.053-.212.062-.209.062-.183.06-.156.14-.315c.174-.358.389-.691.639-.995l.205-.234.205-.207.192-.175.234-.189.25-.178a5.014 5.014 0 011.265-.607l.216-.063.22-.054.412-.073.154-.018.105-.009C16.73 6.006 16.866 6 17 6l.25.006.285.022.28.038.214.04.261.062.273.081c.245.08.483.18.71.296l.23.125.21.129c.078.05.154.103.23.158l.206.159.199.17.209.2.28.308a4.978 4.978 0 011.158 2.99l.005.209v.298a7.005 7.005 0 00-4.417 3.909 2 2 0 003.598 1.742l.162-.338a3 3 0 01.656-.84v6.762l-.004.166a3.5 3.5 0 01-5.223 2.853l-.332-.202a2 2 0 00-2.058 3.426A7.473 7.473 0 0018.5 30a7.468 7.468 0 003.5-.865V37l-.006.253a4.977 4.977 0 01-1.104 2.888l-.163.19-.189.201c-.13.13-.266.253-.407.367l-.042.033-.22.165-.14.094-.265.16c-.15.085-.304.163-.462.232l-.014.007-.205.082-.18.065c-.12.041-.243.078-.368.11l-.152.035-.107.022-.286.049-.113.013-.277.025c-.1.006-.2.009-.3.009l-.127-.002a5.051 5.051 0 01-.434-.03l-.078-.009-.213-.032-.224-.044-.275-.069-.224-.068a5.002 5.002 0 01-.78-.335l-.051-.027-.09-.05-.273-.17A4.99 4.99 0 0112 37l.006-.27.032-.416a2 2 0 00-3.184-1.672A6.995 6.995 0 016 29c0-1.59.527-3.09 1.486-4.313l.63-.804a2 2 0 00.397-3.913l-.364-.092a3.025 3.025 0 01-1.517-1.036l-.122-.169-.089-.139-.092-.168.014.03-.034-.067-.056-.123a2.989 2.989 0 01-.247-1.014L6 17c0-.062.002-.123.006-.184l.006-.084.009-.088c.042-.354.146-.69.301-.997l.098-.179.094-.148.101-.14.094-.118.196-.21.114-.106.143-.118.084-.063.166-.11c.094-.06.19-.113.29-.16l.145-.066-.032.013c.051-.022.103-.043.156-.062l.083-.03.085-.026-.131.044.15-.048zM30.59 6.017l.133-.01L31 6l.225.005.178.011.226.023.235.036.293.06a4.973 4.973 0 011.71.769l.112.081.227.178.206.182.2.199A4.992 4.992 0 0136 11a4.995 4.995 0 01-2.408 4.277 2 2 0 102.077 3.418 8.996 8.996 0 003.803-4.658 2.979 2.979 0 011.53.728l.136.13.156.172.137.174.092.136.07.115.123.235a3.004 3.004 0 01.232 1.831l-.038.173-.046.161-.074.21a3.025 3.025 0 01-.421.74l-.124.149a2.986 2.986 0 01-1.613.943l-.145.036a2 2 0 00.395 3.913l.633.806A6.947 6.947 0 0142 29a6.987 6.987 0 01-2.856 5.641 2 2 0 00-3.172 1.823l.02.267L36 37c0 .73-.157 1.424-.438 2.05l-.119.245a5.02 5.02 0 01-1.163 1.48l-.193.157-.226.169a4.977 4.977 0 01-1.347.665l-.202.06-.279.067-.182.035-.302.042-.321.025-.289.005-.239-.009-.284-.025-.11-.014-.237-.038-.148-.031-.221-.054a4.931 4.931 0 01-.379-.118l-.224-.085-.101-.043-.307-.148a5.02 5.02 0 01-.466-.276l-.175-.123-.177-.136-.197-.167-.112-.102-.101-.099-.114-.118-.205-.233A4.98 4.98 0 0126 37v-7.865a7.467 7.467 0 003.5.865 7.473 7.473 0 004.118-1.231 2 2 0 10-2.198-3.342l-.192.118a3.5 3.5 0 01-5.223-2.853l-.004-.166v-6.762c.27.24.496.53.666.86l.152.318a2 2 0 003.598-1.742A7.006 7.006 0 0026 11.29V11l.005-.217a4.977 4.977 0 011.11-2.93l.159-.185c.06-.068.122-.134.185-.197l.224-.212.308-.254.182-.13.219-.143.157-.091.22-.117c.299-.15.615-.27.943-.357l.144-.036.106-.023.343-.06.285-.031z" fill="currentColor"/></svg>
              <span>AI助手</span>
            </div>

          </div>

          <div v-if="data.files[data.currentFile.index-1].system.search.show">

            <select v-model="data.files[data.currentFile.index-1].system.search.key" class="opBarItem" style="border: 1px solid rgb(230,230,230);position: relative;top: -4px;background-color: #fff;min-width: 100px;width: auto">
              <option v-for="item in data.files[data.currentFile.index-1].system.search.selectKeys" :value="item.key">{{ item.name }}</option>
            </select>

            <input placeholder="请输入关键字" style="height: 28px;border: 0;line-height: 28px;position: relative;top: -4px;border: 1px solid rgb(230,230,230);background-color: #fff;" v-model="data.files[data.currentFile.index-1].system.search.text" class="opBarItem" type="text" name="" id="">

            <div class="opBarItem" @click="()=>{
                let searchInput = data.files[data.currentFile.index-1].system.search.text
                if(searchInput.length==0){
                  return
                }
                let key = data.files[data.currentFile.index-1].system.search.key
                let currentColumns = Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)
                let searchColIndex = 0
                for(let i=0;i<currentColumns.length;i++){
                  if(currentColumns[i]==key){
                    searchColIndex = i
                    break
                  }
                }
                let tableTemp = document.getElementById('mainTable').getElementsByTagName('tr')
                let table = []
                for(let i=1;i<tableTemp.length;i++){
                  table.push(tableTemp[i].getElementsByTagName('td'))
                }
                let res = []
                for(let i=0;i<table.length;i++){
                  if(table[i][searchColIndex].innerText.indexOf(searchInput)!=-1){
                    res.push(i)
                  }
                }
                data.files[data.currentFile.index-1].system.search.results = res
                data.files[data.currentFile.index-1].system.search.showResult = true
              }">
              <svg style="position: relative;left: 2px" width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M22 5c9.389 0 17 7.611 17 17 0 3.549-1.087 6.844-2.947 9.57l6.782 6.782a1 1 0 010 1.414l-1.697 1.698a1 1 0 01-1.414 0l-6.604-6.605A16.934 16.934 0 0122 39c-9.389 0-17-7.611-17-17S12.611 5 22 5zm0 4.2C14.93 9.2 9.2 14.93 9.2 22S14.93 34.8 22 34.8 34.8 29.07 34.8 22 29.07 9.2 22 9.2z" fill="currentColor"/></svg>
              <span></span>
            </div>

            <div class="opBarItem" @click="()=>{
                data.files[data.currentFile.index-1].system.search.show = false
                data.files[data.currentFile.index-1].system.search.text = ''
                data.files[data.currentFile.index-1].system.search.showResult = false
              }">
              <svg width="16" height="16" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
              <span>关闭搜索</span>
            </div>
          </div>

        </div>

        <div style="width: 100%;height: 48px;position: fixed;left: 0;top: calc(48px + 72px);background-color: rgba(255,255,255,0.6);z-index: 900;backdrop-filter: blur(10px);border-bottom: 1px solid rgb(230,230,230);border-top: 1px solid rgb(230,230,230);transition: top .3s cubic-bezier(0.23, 1, 0.320, 1);text-align: center;top: calc(48px + 72px)"
             :style="{
            top:data.tableOnScrollTopBar.show?'calc(48px + 72px)':'calc(72px)',
            opacity:showTableOnScrollTop()==false?0:1
          }"
        >
          <div v-if="data.showTableOnScrollTopAndBottomSum"
               :style="{
              width: getTableWidth(),
              transform: 'translateX(' + (data.tableWrapperScrollLeft*-1).toString() + 'px)'
            }"
               style="height: 36px;display: inline-block;">
            <div style="font-size: 12px;color: gray;position: relative;top:3px;max-width: 100vw"
                 :style="{
                    transform:'translateX(' + data.tableWrapperScrollLeft.toString() + 'px)'
                  }"
            >
              {{ data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].title }}
            </div>
            <div>

              <div v-for="(item,index) in getBottomContent(false)" style="display: inline-block;font-size: 14px;height: 36px;position: relative;margin-left: 0.9px;margin-right: 0.9px;margin-top: 2px;border-radius: 5px;margin-top: 5px"
                   :style="{
                  width:item.width.toString() + 'px'
              }"
              >
                <div @mouseup="()=>{
                  if(index==0){
                    showRightMenu('tableIndexTitleCell',{
                      columnKey:'KpAF0'
                    })
                    return
                  }
                  showRightMenu('tableTitleCell',{
                      columnIndex:index,
                      columnKey:item.key
                    })
                }" class="rightMenuButton" style="width: 100%;height: 100%;line-height: 20px;position: absolute;overflow: hidden;padding-left: 5px;padding-right: 5px;padding-top: 0;padding-bottom: 0;height: 24px;border-radius: 5px;padding-top: 2px" :title="item.name">
                  <span>{{ item.name }}</span>
                </div>
              </div>

            </div>
          </div>
        </div>

        <div id="tableScrollPage" style="width: 100%;height: calc(100% - 108px);overflow: auto;text-align: center;padding-bottom: 20px;position: relative">
          <div style="width: calc(100vw - 20px);margin-left: 10px;">
            <div style="padding-top: 20px;display: inline-block;position: relative;">

              <div id="tableTitle" style="width: 100%;"

                   @click="()=>{
                showInputForm('编辑标题','请输入标题内容',()=>{
                  operationStackAppend('editTitle',data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].title)
                  data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].title = data.input.text
                  controlForm('input',0)
                  setCurrentFileUnsave()
                },data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].title)
              }"
                   :style="{
                   color:(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].title==''?'lightGray':'black'),

                 }"
              >
                <div
                    class="mcell" style="max-width: calc(100vw - 20px);border-radius: 5px;box-shadow: 0 1px 3px rgba(0,0,0,0.2);padding: 20px 10px;font-size: 24px;padding-bottom: 15px;"
                    :style="{
                    transform:'translateX(' + data.tableWrapperScrollLeft.toString() + 'px)'
                  }"
                >
                  <p style="word-break: break-all">{{ data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].title==''?'请输入标题':data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].title }}</p>

                  <p style="font-size: 12px;color: gray;margin-top: 5px">上次编辑于{{ data.files[data.currentFile.index-1].system.recentEditTime }}</p>
                </div>

              </div>

              <div class="mcell" style="background-color: white;text-align: center">
                <table id="mainTable" style="table-layout: auto;border-radius: 5px;border: solid transparent;margin: 0 auto;user-select: unset !important" border="1px">
                  <tr id="mainTableTitles">
                    <td @mouseup="()=>{
                    if(item=='KpAF0'){
                      showRightMenu('tableIndexTitleCell',{
                        columnKey:'KpAF0'
                      })
                      return
                    }
                    showRightMenu('tableTitleCell',{
                      columnIndex:index,
                      columnKey:item
                    })
                  }"
                        :style="{
                            minWidth:(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[item].minWidth==undefined || data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[item].minWidth=='' || data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[item].minWidth==0)?'unset':(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[item].minWidth.toString() + 'px')
                        }"
                        style="border-radius: 5px;" class="table-line" v-for="(item,index) in Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)">{{ data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[item].name }}</td>
                  </tr>
                  <tr
                      @click="()=>{
                  let cc = Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)
            let cd = data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data[index]
            let re = []
            for(let i=1;i<cc.length;i++){
              if(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].type!='text' && data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].type!='number' && data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].type!='date')continue
              re.push({
                keyName:cc[i],
                value:cd[cc[i]],
                type:data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].type,
                name:data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns[cc[i]].name
              })
            }
            data.addRow.inputValues = re
            data.addRow.isEdit = true
            data.rightMenu.data = {
              dataIndex:index
            }
            controlForm('addRow',1)
                }" class="table-line" style="border: 1px;" v-for="(item,index) in data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].data">
                    <td :class="isSearchResult(index)?['table-line-selected']:['table-line']"
                        @mouseup="(e)=>{
                    let mpy = 0
                    if(data.rightMenu.topT + 405>data.window.height){
                      mpy = -405
                    }
                    if(data.rightMenu.topT+mpy<48){
                      mpy = 48 - data.rightMenu.topT
                    }
                    showRightMenu('tableCell',{
                      dataIndex:index,
                      columnKey:item2,
                      columnIndex:index2
                    },0,mpy)
                  }" style="border-radius: 5px;border: 1px solid rgb(230,230,230);" v-for="(item2,index2) in Object.keys(data.files[data.currentFile.index-1].data.sheets[data.files[data.currentFile.index-1].system.currentSheetIndex-1].columns)">
                      {{ tableCellGetValue(item,item2) }}</td>
                  </tr>
                </table>
              </div>

            </div>
          </div>


        </div>

        <div v-if="data.showTableOnScrollTopAndBottomSum==true" style="width: 100%;height: 36px;position: fixed;box-shadow: 0 0 3px rgba(0,0,0,0.2);left: 0;bottom: 36px;background-color: #fff;z-index: 1000;border-bottom: 1px solid rgb(230,230,230);text-align: center">
          <div v-if="document.getElementById('mainTableTitles')!=null" style="height: 36px;display: inline-block;position: relative"
               :style="{
              width: getTableWidth(),
              transform: 'translateX(' + (data.tableWrapperScrollLeft*-1).toString() + 'px)'
            }"
          >
            <div v-for="(item,index) in getBottomContent()" style="display: inline-block;font-size: 14px;line-height: 31px;border: 1px solid rgb(230,230,230);height:31px;position: relative;margin-left: 0.9px;margin-right: 0.9px;margin-top: 2px;border-radius: 5px"
                 :style="{
                  width:item.width.toString() + 'px',
                  border:item.text==''?'1px solid rgba(230,230,230,0)':'1px solid rgb(230,230,230)'
              }"
            >
              <div @mouseup="()=>{
                  if(index==0){
                    showRightMenu('tableIndexTitleCell',{
                      columnKey:'KpAF0'
                    },0,-65)
                    return
                  }
                  showRightMenu('tableTitleCell',{
                      columnIndex:index,
                      columnKey:item.key
                    },0,-182)
                }" style="width: 100%;height: 100%;line-height: 31px;position: absolute;overflow: hidden;padding-left: 0px;padding-right: 0px">
                {{ item.text }}
              </div>
            </div>
          </div>
        </div>

        <div style="width: 100%;height: 36px;position: absolute;left: 0;bottom: 0;padding-left: 72px;padding-right: 36px;box-shadow: 0 0 3px rgba(0,0,0,0);z-index: 1001;background-color: #fff;">

          <div v-if="data.files[data.currentFile.index-1].system.search.show==false" style="width: 36px;height: 36px;display: inline-block;min-width: 0;padding: 0 0;position: absolute;right: 0px;top: 0;z-index: 1000" class="fileBar-fileItem"
               :style="{
            display:sheetBarLetfRightVisible(0)?'unset':'none'
            }"
               @click="sheetBarLeftRightClick(0)"
          >
            <svg style="position: absolute;left: 9px;top: 7px;scale: 0.85;" width="20" height="20" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M32.556 24L16.293 7.737a1 1 0 010-1.415l1.414-1.414a1 1 0 011.414 0L36.8 22.586a2 2 0 010 2.828L19.121 43.092a1 1 0 01-1.414 0l-1.414-1.414a1 1 0 010-1.414L32.556 24z" fill="currentColor"/></svg>
          </div>

          <div v-if="data.files[data.currentFile.index-1].system.search.show==false" style="width: 36px;height: 36px;display: inline-block;min-width: 0;padding: 0 0;position: absolute;left: 36px;top: 0;z-index: 1000;" class="fileBar-fileItem"
               :style="{
            display:sheetBarLetfRightVisible(1)?'unset':'none'
            }"
               @click="sheetBarLeftRightClick(1)"
          >
            <svg style="position: absolute;left: 9px;top: 7px;scale: 0.85" width="20" height="20" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M15.272 24l16.263 16.264a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L11.03 25.414a2 2 0 010-2.828L28.706 4.908a1 1 0 011.414 0l1.415 1.414a1 1 0 010 1.415L15.271 24z" fill="currentColor"/></svg>
          </div>

          <div v-if="data.files[data.currentFile.index-1].system.search.show==false" title="新建表格" @click="data.new.newFileName = '';controlForm('newSheet',1)" style="width: 36px;height: 36px;display: inline-block;min-width: 0;padding: 0 0;position: absolute;left: 0px;top: 0;z-index: 1000" class="fileBar-fileItem">
            <svg style="position: absolute;left: 9px;top: 7px;scale: 0.85" width="20" height="20" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M24.718 4c.446 0 .608.046.77.134.163.087.291.215.378.378.088.163.134.324.134.77V22h16.718c.446 0 .607.046.77.134.163.087.291.215.378.378.088.162.134.324.134.77v1.436c0 .446-.046.608-.134.77a.908.908 0 01-.378.378c-.163.088-.324.134-.77.134H26L26 42.718c0 .446-.046.607-.134.77a.908.908 0 01-.378.378c-.162.088-.324.134-.77.134h-1.436c-.446 0-.608-.046-.77-.134a.908.908 0 01-.378-.378c-.088-.163-.134-.324-.134-.77L21.999 26 5.282 26c-.446 0-.607-.046-.77-.134a.908.908 0 01-.378-.378c-.088-.162-.134-.324-.134-.77v-1.436c0-.446.046-.608.134-.77a.908.908 0 01.378-.378c.163-.088.324-.134.77-.134L22 21.999V5.282c0-.446.046-.607.134-.77a.908.908 0 01.378-.378c.162-.088.324-.134.77-.134h1.436z" fill="currentColor"/></svg>
          </div>

          <div v-if="data.files[data.currentFile.index-1].system.search.show==false" id="sheetBarWrapper" style="height: 36px;width: 100%;position: relative">
            <div @mouseup="showRightMenu('sheetBar',undefined,0,-40)" style="width: 100%;height: 36px;position: absolute;">

            </div>

            <div style="width: 10000px;position: relative">
              <div v-if="data.files.length!=0" id="sheetBar" style="display: inline-block;position: absolute;top: 0;"
                   :style="{
                left:data.sheetBar.left.toString() + 'px',
                transition:data.sheetBar.showAnimation?'left .5s cubic-bezier(0.23, 1, 0.320, 1)':'unset'
              }"
              >

                <div :title="item.name" v-for="item in data.files[data.currentFile.index-1].data.sheets" class="fileBar-fileItem"
                     @click="sheetBarItemClick(item.name)"
                     @mouseup="showRightMenu('sheetBarItem',item,0,-157)"
                     style="display: inline-block;padding-right: 10px;min-width: 0"
                     :class="(item.name == data.files[data.currentFile.index-1].system.currentSheetName)?['fileBar-fileItem fileBar-fileItem-selected']:['fileBar-fileItem']"
                >
                  {{ item.name }}
                </div>

              </div>
            </div>



          </div>

          <div v-if="data.files[data.currentFile.index-1].system.search.show" style="line-height: 36px;font-size: 14px">
            <span style="margin-right: 10px;color: gray">搜索时不能切换和新建工作表。</span><span style="text-decoration: underline;color: rgb(0,120,215);cursor: pointer;margin-right: 20px" @click="()=>{
              data.files[data.currentFile.index-1].system.search.show = false
              data.files[data.currentFile.index-1].system.search.text = ''
                data.files[data.currentFile.index-1].system.search.showResult = false
            }">关闭搜索</span><span v-if="data.files[data.currentFile.index-1].system.search.showResult" style="border-radius: 5px;background-color: rgb(0,120,215);color: white;padding: 0 5px">共{{ data.files[data.currentFile.index-1].system.search.results.length }}个搜索结果</span>
          </div>

        </div>

      </div>

    </div>


    <div v-if="currentFile" style="width: 500px;height: calc(100vh - 265px);position: fixed;top: 175px;right: 15px;background-color: rgb(242,242,242,0.3);border-radius: 10px;z-index: 60000;box-shadow: 0 4px 10px rgba(0,0,0,0.2);backdrop-filter: blur(20px);border: 1px solid lightgray;transition-property: right;transition-timing-function: cubic-bezier(0.23, 1, 0.320, 1);transition-duration: .3s"
      :style="{
        right:data.aiAides.show?'15px':'-510px',
        transitionProperty:data.aiAides.showAnimation?'right':'0'
      }"
    >

      <div style="padding-left: 30px;padding-top: 34px;padding-right: 30px;padding-bottom: 10px;position: relative">
        <p style="font-size: 22px;font-weight: bold">AI助手</p>
        <p style="font-size: 12px;color: rgba(0,0,0,0.6)">对话内容由AI生成，请注意甄别</p>
        <svg v-if="data.aiAides.working==false" @click="async ()=>{
          data.aiAides.showAnimation = true
          data.aiAides.show = false
          await tool.waitSeconds(0.5)
          data.aiAides.showAnimation = false
        }" style="position: absolute;right: 30px;top: 34px;cursor: pointer;" width="24" height="24" viewBox="0 0 48 48" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M23.886 20.957L37.32 7.522a1 1 0 011.414 0l1.414 1.414a1 1 0 010 1.415L26.714 23.786 40.149 37.22a1 1 0 010 1.414l-1.414 1.414a1 1 0 01-1.414 0L23.886 26.614 10.45 40.049a1 1 0 01-1.415 0l-1.414-1.414a1 1 0 010-1.414l13.435-13.435L7.622 10.35a1 1 0 010-1.415l1.414-1.414a1 1 0 011.415 0l13.435 13.435z" fill="currentColor"/></svg>
      </div>

      <div v-if="data.aiAides.show" style="width: 100%;height: calc(100vh - 450px);overflow-y: scroll;padding: 10px 30px" id="aichat">

        <chat-item @endLoading="endLoading" @disableAnimation="disableAnimation" v-for="(item,index) in data.aiAides.contents" :config="item"></chat-item>

      </div>

      <div v-if="data.aiAides.show" style="width: 100%;height: 93px;position: relative;padding: 10px 30px;">
        <textarea @keydown="(e)=>{
          if(e.code=='Enter'){
            sendMessage()
          }
        }" v-model="data.aiAides.input" :disabled="data.aiAides.working" :placeholder="data.aiAides.working?(data.aiAides.thinking?'正在思考...':'正在回复...'):'输入问题'" style="max-width: calc(100% - 60px);min-width: calc(100% - 60px);border: 1px solid lightgrey;background-color: white;border-radius: 5px;min-height: 65px;max-height: 65px;padding: 10px 10px;font-size: 14px;font-family: 微软雅黑" name="" id="" cols="30" rows="10"></textarea>
        <button v-if="data.aiAides.working==false" style="margin-top: 0;margin-right: 0;position: absolute;right: 30px;top: 10px;padding-left: 10px;padding-right: 10px" @click="sendMessage">发送</button>
        <div v-if="data.aiAides.working==true" style="display: inline-block;margin-top: 0;margin-right: 0;position: absolute;right: 20px;top: 10px;padding-left: 10px;padding-right: 10px">
          <div style="--size: 44px; --dot-size: 5px; --dot-count: 6; --color: black; --speed: 1s; --spread: 60deg;display: inline-block;" class="dots">
            <div style="--i: 0;" class="dot"></div>
            <div style="--i: 1;" class="dot"></div>
            <div style="--i: 2;" class="dot"></div>
            <div style="--i: 3;" class="dot"></div>
            <div style="--i: 4;" class="dot"></div>
            <div style="--i: 5;" class="dot"></div>
          </div>
        </div>

      </div>

    </div>

  </div>
</template>

<script>
import {deleteArrElemByIndex} from "@/utlis/tool";
const os = require('os');
const { remote,app,clipboard } = require("electron");
const tool = require('./utlis/tool')
const path = require('path')
import selectFolder from "./components/selectFolder.vue";
import {appendFile} from "fs";
const XLSX = require('xlsx')
const { dialog } = require('electron').remote
const fs = require('fs')
const axios = require('axios')
import chatItem from "@/components/chatItem.vue";


  export default {
    computed: {
      tool() {
        return tool
      },
      clipboard() {
        return clipboard
      },
      document(){
        return document
      },
      currentFile(){
        return this.data.files[this.data.currentFile.index-1]
      }
    },
    components:{
      selectFolder,chatItem
    },

    data() {
      return {
        data:{
          aiAides: {
            showAnimation:false,
            input:'',
            working:false,
            show:false,
            contents:[],
            used:false,
            thinking:false,
          },
          appStarted:false,
          showStartPage:2,
          cwm:false,
          files:[

          ],
          currentFile:{
            index:0,
            path:''
          },
          setting:{
            readData:undefined,
            drives:[],
            form:{
              saveDrive:'',
              closeAllFiles:''
            }
          },
          new:{
            currentFolderPath:'',
            currentFolderFileNames:[],
            newFolderName:'',
            showSelectFolder:true,
            newFileName:'',
            defaultPath:''
          },
          warning:{
            title:'',
            content:'',
            callBack:undefined,
            showCancel:false,
            buttonTexts:{
              cancel: '取消',
              ok:'确定',
              secondary: undefined
            },
            showSecondary:false
          },
          fileList:{
            all:[],
            current:[],
            path:[]
          },
          rightMenu:{
            name:'',
            left:0,
            top:0,
            show:false,
            data:undefined,
            leftT:0,
            topT:0,
            allowShow:false,
          },
          fileBar:{
            left:0,
            showAnimation:false,
          },
          sheetBar:{
            left:0,
            showAnimation:false,
          },
          input:{
            text:'',
            title:'',
            subTitle:'',
            callBack:undefined
          },
          addCol:{
            name:'',
            type:'text',
            sumType:'none',
            selectCol:{

            },
            isEdit:false,
          },
          addRow:{
            isEdit:false,
            inputValues:[

            ]
          },
          window:{
            width:0,
            height:0
          },
          editCol:{
            keyName:'',
            name:'',
            sumType:'',
            type:'',
            isIndexCol:false,
            minWidth:'',
          },
          tableOnScrollTopBar:{
            show:false,
            titles:[]
          },
          showTableOnScrollTopAndBottomSum:true,
          tableWrapperScrollLeft:0,
          sort:{
            columns:[],
            selectColumn:'',
            isRe:false,
          },
          themes:{
            'default':{
              titleBarColor:'rgb(0,0,0)',
              titleBarBackground:'rgb(242,242,242)',
              controlButtonColor:'',
              name:'默认'
            },
            'white':{
              titleBarColor:'rgb(0,0,0)',
              titleBarBackground:'rgb(255,255,255)',
              controlButtonColor:'',
              name:'白色'
            },
            'blue':{
              titleBarColor:'white',
              titleBarBackground:'rgb(0,120,215)',
              controlButtonColor:'_dark',
              name:'蓝色'
            },
          },
          backgroundImage:'url("./assets/trans.png")',
          import:{
            isDir:false,
          }
        }
      }
    },


    methods: {
      endLoading(){
        this.data.aiAides.working = false
      },
      getTable(){
        let re = '' + this.currentFile.data.sheets[this.currentFile.system.currentSheetIndex-1].name + '\n'
        re += '' + this.currentFile.data.sheets[this.currentFile.system.currentSheetIndex-1].title + '\n'
        let tablet = document.getElementById('mainTable').getElementsByTagName('tr')
        for(let i=0;i<tablet.length;i++){
          let t = tablet[i].getElementsByTagName('td')
          for(let i=0;i<t.length;i++){
            re += t[i].innerText + '\t'
          }
          re += '\n'
        }
        return re + '\n'
      },
      async sendMessage(){
        if(this.data.aiAides.working || this.data.aiAides.input==''){
          return
        }
        this.data.aiAides.working = true
        this.data.aiAides.thinking = true
        this.data.aiAides.contents.push({
          position:1,
          text:this.data.aiAides.input,
          showAnimation:false,
        })
        setTimeout(()=>{
          try {
            document.getElementById('aichat').scrollTop = 10000000000
          }catch(e){

          }
        },10)
        var data = this.getTable() + this.data.aiAides.input
        this.data.aiAides.input = ''
        await tool.waitSeconds(0.5)
        this.data.aiAides.input = ''
        axios.post('/chat', data, {
          headers: {
            'Content-Type': 'application/json' // 设置请求头部
          }
        })
            .then(async (response) => {
              this.data.aiAides.thinking = false
              let re = response.data
              let tt = {
                position:-1,
                text:re,
                showAnimation:false,
              }
              this.data.setting.readData.aiAidesContents = JSON.parse(JSON.stringify(this.data.aiAides.contents))
              this.data.setting.readData.aiAidesContents.push(JSON.parse(JSON.stringify(tt)))
              tt.showAnimation = true
              this.data.aiAides.contents.push(tt)
              this.data.aiAides.used = true
              await tool.updateFileContent(this.getAppPath() + "\\ExcelEasyConfig.JSON",JSON.stringify(this.data.setting.readData))
            })
            .catch(async (error) => {
              this.data.aiAides.working = false
              let re = '很抱歉，我暂时不能回答此问题，请稍后再试。'
              let tt = {
                position:-1,
                text:error.toString(),
                showAnimation:false,
              }
              this.data.setting.readData.aiAidesContents = JSON.parse(JSON.stringify(this.data.aiAides.contents))
              this.data.setting.readData.aiAidesContents.push(JSON.parse(JSON.stringify(tt)))
              tt.showAnimation = true
              this.data.aiAides.contents.push(tt)
              this.data.aiAides.used = true
              await tool.updateFileContent(this.getAppPath() + "\\ExcelEasyConfig.JSON",JSON.stringify(this.data.setting.readData))
              console.log(error);
            });
      },
      disableAnimation(){
        for(let i=0;i<this.data.aiAides.contents.length;i++){
          this.data.aiAides.contents[i].showAnimation = false
        }
      },
      async opendedAiAds(){
        if(this.data.aiAides.used==false){
          this.data.aiAides.contents.push({
            position:-1,
            text:'您好，我是您的个人助理，我可以帮您分析当前表格里的数据。',
            showAnimation:true,
          })
          this.data.aiAides.used=true
        }
        try {
          document.getElementById('aichat').scrollTop = 10000000000
        }catch(e){

        }

      },
      async getBackground(){
        axios.get('https://api.xygeng.cn/bing/').then(res=>{
          this.data.backgroundImage = 'url(' + res.data.data.url + ') !important'
        }).catch(err=>{
          console.log(err)
        })
      },
      appendFile,
      async startImportFile(saveToPath){
        let currentFile = await this.readExcelFile(saveToPath.path)
        let cf = []
        for(let i=0;i<currentFile.length;i++){
          let re = this.modifyImportData(currentFile[i])
          if(re){
            cf.push(re)
          }
        }
        if(cf.length==0){
          this.setWarningFormTitleAndContentAndShowForm('导入失败','此表格没有格式符合要求的工作表。',false,()=>{})
          return
        }
        let fileData = {
          name: '导入 ' + ((new Date).getFullYear()).toString() + '_' + ((new Date).getMonth()+1).toString() + '_' + ((new Date).getDate()).toString() + ' ' + ((new Date).getHours()).toString() + '_' + ((new Date).getMinutes()).toString() + '_' + ((new Date).getSeconds()).toString() + ' ' + tool.removeFileNameExt(saveToPath.fileName) ,
          sheets: cf
        }
        // let fp = this.data.setting.readData.saveDrive + "\\ExcelEasyData\\" + fileData.name + ".JSON"
        let fp = this.data.setting.readData.saveDrive + "\\ExcelEasyData\\" + this.getCurrentFileListFolderPath() + '\\' + fileData.name + ".JSON"
        await tool.createFileWithContent(fp,JSON.stringify(fileData))
        await tool.waitSeconds(1)
        await this.refreshFileList()
      },
      async importFile(){
        this.setWarningFormTitleAndContentAndShowForm('导入','请选择格式为.xlsx或.xls的文件或包含此类文件的文件夹，文件中每个工作表的首行为表的标题，第二行为表的每列的标题。',true,async ()=>{
          let saveToPath = await tool.selectFile()
          if(saveToPath==false){
            return
          }
          this.showLoading()
          await this.startImportFile(saveToPath)
          this.hideLoading()
        },false,{
          ok:'选择文件',
          cancel: '取消',
          secondary: '选择文件夹'
        },async ()=>{
          let saveToPath = await tool.chooseFolder()
          console.log(saveToPath)
        })
      },
      modifyImportData(sheetData){
        if(sheetData.data.length<2){
          return false
        }
        sheetData['title'] = sheetData.data[0][0]
        let columns = {
          'KpAF0':{
            name:'序号',
            type:'number',
            sumType:'none',
          }
        }
        let b64s = new tool.Base64String(1)
        for(let i=0;i<sheetData.data[1].length;i++){
          columns[b64s.get()] = {
            name:sheetData.data[1][i],
            type:(this.isColAllNumber(sheetData.data.slice(2),i)?'number':'text'),
            sumType:'none',
          }
          b64s.add1()
        }
        sheetData['columns'] = columns
        let data = []
        for(let i=2;i<sheetData.data.length;i++){
          let cdata = {
            'KpAF0':i-1,
          }
          for(let i1=0;i1<sheetData.data[1].length;i1++){
            if(sheetData.data[i][i1]==null || sheetData.data[i][i1]==''){
              if(sheetData.columns[Object.keys(sheetData.columns)[i1+1]].type=='text'){
                sheetData.data[i][i1] = ''
              }else{
                sheetData.data[i][i1] = '0'
              }
            }
            cdata[Object.keys(sheetData.columns)[i1+1]] = sheetData.data[i][i1]
          }
          data.push(cdata)
        }
        sheetData.data = data
        sheetData.config = {}
        return sheetData
      },
      isColAllNumber(arr,index){
        for(let i=0;i<arr.length;i++){
          if(arr[i][index]==''){
            continue
          }
          try {
            if(typeof(JSON.parse(arr[i][index]))!='number'){
              return false
            }
          }catch (e){
            return false
          }
        }
        return true
      },
      readExcelFile(filePath) {
        return new Promise((resolve, reject) => {
          const xhr = new XMLHttpRequest();
          xhr.open('GET', filePath, true);
          xhr.responseType = 'arraybuffer';
          xhr.onload = function (e) {
            const arraybuffer = xhr.response;
            const data = new Uint8Array(arraybuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheets = [];
            workbook.SheetNames.forEach((sheetName) => {
              const worksheet = workbook.Sheets[sheetName];
              const sheetJson = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
              const sheetData = { name: sheetName, data: sheetJson };
              sheets.push(sheetData);
            });
            resolve(sheets);
          };
          xhr.onerror = function (e) {
            reject('无法读取Excel文件');
          };
          xhr.send();
        });
      },
      refreshTopAndBottom(){
        this.data.showTableOnScrollTopAndBottomSum = false
        setTimeout(()=>{
          this.data.showTableOnScrollTopAndBottomSum = true
        },1)
      },
      async exportExcelFileFromPath(filePath){
        let fileData = JSON.parse(await tool.readFile(filePath))
        // let saveToPath = await this.selectFolder()
        this.showLoading()

        const workbook = XLSX.utils.book_new()
        for(let i=0;i<fileData.sheets.length;i++){
          let re = [[fileData.sheets[i].title]]
          let currentSheet = fileData.sheets[i]
          let titles = []
          for(let i1=0;i1<Object.keys(currentSheet.columns).length;i1++){
            titles.push(currentSheet.columns[Object.keys(currentSheet.columns)[i1]].name)
          }
          // re.push(titles)
          
          let tablet = this.document.getElementById('mainTable').getElementsByTagName('tr')
          let allNum = tablet.length-1
          for(let i=0;i<tablet.length;i++){
            let re1 = tablet[i].getElementsByTagName('td')
            let res = []
            for(let i1=0;i1<re1.length;i1++){
              res.push(re1[i1].innerText)
            }
            re.push(res)
          }

          let colSums = []
          let colSumsT = this.getBottomContent()
          let flag = true
          if(colSumsT[0].text=='')flag = false
          for(let i=0;i<colSumsT.length;i++){
            colSums.push(colSumsT[i].text)
          }
          colSums[0] = '合计'
          re.push(colSums)
          if(flag){
            re.push(['共' + allNum.toString() + '个记录'])
          }
          const worksheet = XLSX.utils.aoa_to_sheet(re) 
          const mergeRange = { s: { r: 0, c: 0 }, e: { r: 0, c: Object.keys(currentSheet.columns).length-1 } }
          worksheet['!merges'] = [mergeRange]
          XLSX.utils.book_append_sheet(workbook, worksheet, currentSheet.name)
        }
        XLSX.writeFile(workbook, fileData.name + ".xlsx")
        this.hideLoading()
      },
      selectFolder() {
  return new Promise((resolve, reject) => {
    dialog.showOpenDialog({
      properties: ['openDirectory']
    }).then(result => {
      if (result.canceled) {
        resolve(false); // 用户取消了文件夹选择
      } else {
        resolve(result.filePaths[0]); // 返回选择的文件夹路径
      }
    }).catch(err => {
      reject(err); // 处理任何错误
    });
  });
},
      showTableOnScrollTop(){
        try {
          return (document.getElementById('tableScrollPage').scrollTop>=document.getElementById('mainTable').offsetTop)
        }catch (e){
          return false
        }
      },
      isSearchResult(index){
        if(this.data.files[this.data.currentFile.index-1].system.search.showResult==false){
          return false
        }
        for(let i=0;i<this.data.files[this.data.currentFile.index-1].system.search.results.length;i++){
          if(this.data.files[this.data.currentFile.index-1].system.search.results[i]==index){

            return true
          }
        }
        return false
      },
      getColSumContent(columnIndex,showUnit){
        let sumType = this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[columnIndex]].sumType
        let type = this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[columnIndex]].type
        if(sumType=='none')return ''
        let table1 = document.getElementById('mainTable').getElementsByTagName('tr')
        let table = []
        for(let i=1;i<table1.length;i++){
          table.push(table1[i].getElementsByTagName('td'))
        }
        if(sumType=='count'){
          let re = 0
          for(let i=0;i<table.length;i++){
            if(table[i][columnIndex].innerText=='' && type=='text'){
              continue
            }
            if(table[i][columnIndex].innerText=='0' && type!='text'){
              continue
            }
            re+=1
          }
          return re.toString() + (showUnit?'个':'')
        }
        if(sumType=='num'){
          return table.length.toString() + (showUnit?'个':'')
        }
        let sum = 0
        for(let i=0;i<table.length;i++){
          sum+=JSON.parse(table[i][columnIndex].innerText)
        }
        if(sumType=='sum'){
          return sum.toString()
        }
        if(sumType=='sumdiv'){
          return (parseInt(sum/(table.length)*100)/100).toString()
        }
        return ''
      },
      getBottomContent(getText = true){
        try {
          let re = undefined
          try {
            re = document.getElementById('mainTableTitles').getElementsByTagName('td')
          }catch(e){
            return []
          }

          let re0 = []
          for(let i=0;i<re.length;i++){
            re0.push({
              width:re[i].clientWidth+1,
              text:getText?this.getColSumContent(i,true):'',
              name:this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i]].name,
              key:Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i]
            })
          }
          if(getText){

          }

          return re0
        }catch (e) {

        }

      },
      getTableWidth(){
        try {
          return (document.getElementById('tableTitle').clientWidth+20).toString() + 'px'
        }catch(e){
          return '0px'
        }
      },
      deleteArrElemByIndex,
      tableCellGetValue(item,item2){
        if(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[item2].type=='sum'){
          return this.tableLineGetSum(item,this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[item2].sumCols,false,this.currentFile.data.sheets[this.currentFile.system.currentSheetIndex-1].columns)
        }
        if(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[item2].type=='sumdiv'){
          return parseInt(this.tableLineGetSum(item,this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[item2].sumCols,true,this.currentFile.data.sheets[this.currentFile.system.currentSheetIndex-1].columns)*10000)/100
        }
        return ' ' + item[item2].toString() + ' '
      },
      tableLineGetSum(lineItem,sumKeys,isDiv,columns){
        let re = 0
        for(let i=0;i<sumKeys.length;i++){
          if(lineItem[sumKeys[i]]==undefined){
            continue
          }
          if(columns[sumKeys[i]].type=='sum'){
            re += this.tableLineGetSum(lineItem,columns[sumKeys[i]].sumCols,false,columns)
          }
          if(columns[sumKeys[i]].type=='sumdiv'){
            re += this.tableLineGetSum(lineItem,columns[sumKeys[i]].sumCols,true,columns)
          }
          if(columns[sumKeys[i]].type=='number'){
            re += JSON.parse(lineItem[sumKeys[i]])
          }
        }
        if(isDiv){
          re /= sumKeys.length
        }
        return re
          //
        // let re = 0
        // for(let i=0;i<sumKeys.length;i++){
        //   if(lineItem[sumKeys[i]]==undefined){
        //     continue
        //   }
        //   re += lineItem[sumKeys[i]]
        // }
        // if(isDiv){
        //   re /= sumKeys.length
        // }
        // return re
      },
      reDo(){
        if(this.data.files[this.data.currentFile.index-1].system.operationStackIndex>=(this.data.files[this.data.currentFile.index-1].system.operationStack.length-1)){
          return
        }
        this.setCurrentFileUnsave()
        this.data.files[this.data.currentFile.index-1].system.operationStackIndex+=1
        let csn = this.data.files[this.data.currentFile.index-1].system.currentSheetName
        this.data.files[this.data.currentFile.index-1].data.sheets = this.data.files[this.data.currentFile.index-1].system.operationStack[this.data.files[this.data.currentFile.index-1].system.operationStackIndex].data
        for(let i=0;i<this.data.files[this.data.currentFile.index-1].data.sheets.length;i++){
          if(this.data.files[this.data.currentFile.index-1].data.sheets[i].name == (this.data.files[this.data.currentFile.index-1].system.operationStack[this.data.files[this.data.currentFile.index-1].system.operationStackIndex].sheetName==undefined?csn:this.data.files[this.data.currentFile.index-1].system.operationStack[this.data.files[this.data.currentFile.index-1].system.operationStackIndex].sheetName)){
            this.data.files[this.data.currentFile.index-1].system.currentSheetIndex = i+1
            this.data.files[this.data.currentFile.index-1].system.currentSheetName = this.data.files[this.data.currentFile.index-1].data.sheets[i].name
            return
          }
        }
      },
      withDrawClick(){
        if(this.data.files[this.data.currentFile.index-1].system.operationStackIndex==0)return
        this.setCurrentFileUnsave()
        if(this.data.files[this.data.currentFile.index-1].system.operationStackIndex==this.data.files[this.data.currentFile.index-1].system.operationStack.length){
          this.data.files[this.data.currentFile.index-1].system.operationStack.push({
            sheetName:undefined,
            data:JSON.parse(JSON.stringify(this.data.files[this.data.currentFile.index-1].data.sheets)),
          })
        }
        this.data.files[this.data.currentFile.index-1].system.operationStackIndex-=1
        let csn = this.data.files[this.data.currentFile.index-1].system.currentSheetName
        this.data.files[this.data.currentFile.index-1].data.sheets = this.data.files[this.data.currentFile.index-1].system.operationStack[this.data.files[this.data.currentFile.index-1].system.operationStackIndex].data
        for(let i=0;i<this.data.files[this.data.currentFile.index-1].data.sheets.length;i++){
          if(this.data.files[this.data.currentFile.index-1].data.sheets[i].name == (this.data.files[this.data.currentFile.index-1].system.operationStack[this.data.files[this.data.currentFile.index-1].system.operationStackIndex].sheetName==undefined?csn:this.data.files[this.data.currentFile.index-1].system.operationStack[this.data.files[this.data.currentFile.index-1].system.operationStackIndex].sheetName)){
            this.data.files[this.data.currentFile.index-1].system.currentSheetIndex = i+1
            this.data.files[this.data.currentFile.index-1].system.currentSheetName = this.data.files[this.data.currentFile.index-1].data.sheets[i].name
            return
          }
        }
      },
      operationStackAppend(){
        for(let i=this.data.files[this.data.currentFile.index-1].system.operationStack.length-1;i>=this.data.files[this.data.currentFile.index-1].system.operationStackIndex;i--){
          this.data.files[this.data.currentFile.index-1].system.operationStack = tool.deleteArrElemByIndex(this.data.files[this.data.currentFile.index-1].system.operationStack,i)
        }
        this.data.files[this.data.currentFile.index-1].system.operationStack.push({
          sheetName:this.data.files[this.data.currentFile.index-1].system.currentSheetName,
          data:JSON.parse(JSON.stringify(this.data.files[this.data.currentFile.index-1].data.sheets)),
        })
        this.data.files[this.data.currentFile.index-1].system.operationStackIndex+=1
      },
      checkIsNumber(string){
        try {
          string = JSON.parse(string)
          if(typeof(string)!='number'){
            return false
          }
        }catch(e){
          return false
        }
        return true
      },
      rowCheckInput(valueInputs){
        function isZS(n){
          return parseInt(n)==n
        }
        for(let i=0;i<valueInputs.length;i++){
          switch (valueInputs[i].type) {
            case 'text':

              break
            case 'number':
              try {
                valueInputs[i].value = JSON.parse(valueInputs[i].value)
                if(typeof(valueInputs[i].value)!='number'){
                  throw new Error()
                }
                if(valueInputs[i].keyName=='KpAF0'){
                  if(isZS(valueInputs[i].value)==false || valueInputs[i].value<=0 || valueInputs[i].value>this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].data.length+1){
                    this.setWarningFormTitleAndContentAndShowForm('提示','第' + (i+1).toString() + '列-' + valueInputs[i].name + '的值应为正整数，并且不大于此表格记录数+1。',false,()=>{})
                    return false
                  }
                }
              }catch (e){
                this.setWarningFormTitleAndContentAndShowForm('提示','第' + (i+1).toString() + '列-' + valueInputs[i].name + '的值应为数字。',false,()=>{})
                return false
              }
              break
            case 'date':

              break
          }
        }
        return true
      },
      addRow_getCurrentSheetCols(){
        let re = []
        for(let i=0;i<Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns).length;i++){
          let type = this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i]].type
          if(type=='sum' || type=='sumdiv'){
            continue
          }
          let defaultValue = undefined
          switch (type) {
            case 'text':
              defaultValue = ''
              break
            case 'number':
              defaultValue = (Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i]=='KpAF0'?(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].data.length+1):0)
              break
            case 'date':
              defaultValue = (new Date).getFullYear().toString() + '-' + (((new Date).getMonth() + 1).toString().length==1?('0' + ((new Date).getMonth() + 1).toString()):(((new Date).getMonth() + 1).toString())) + '-' + ((new Date).getDate().toString().length==1?('0' + (new Date).getDate().toString()):((new Date).getDate().toString()))
              break
          }
          re.push({
            keyName:Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i],
            value:defaultValue,
            type:type,
            name:this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i]].name
          })
        }
        // this.data.addRow.inputValues = re
        return re
      },
      getColumnsNewKey(){
        let keys = Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)
        for(let i=0;i<keys.length;i++){
          let t = new tool.Base64String(0)
          t.setByBase64(keys[i])
          keys[i] = t.getBase10()
        }
        keys.sort()
        let t = new tool.Base64String(keys[keys.length-1])
        t.add1()
        return t.get()
      },
      addCol_getCurrentSheetCols(){
        let re = []
        for(let i=0;i<Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns).length;i++){
          if(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i]].name=='序号')continue
          let type = this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i]].type
          if(type!='number' && type!='sum' && type!='sumdiv'){
            continue
          }
          re.push({
            keyName:Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i],
            name:this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns[Object.keys(this.data.files[this.data.currentFile.index-1].data.sheets[this.data.files[this.data.currentFile.index-1].system.currentSheetIndex-1].columns)[i]].name,
            index:i+1,
            selected:false,
          })
        }
        this.data.addCol.selectCol = re
      },
      showInputForm(title,subTitle,callBack,textDefault=''){
        this.data.input.title = title
        this.data.input.subTitle = subTitle
        this.data.input.callBack = callBack
        this.data.input.text = textDefault
        this.controlForm('input',1)
      },
      sheetBarLetfRightVisible(op){
        if(document.getElementById('sheetBarWrapper')==null || this.data.files.length==0){
          return false
        }
        if(op){
          if(this.data.sheetBar.left>=0){
            return false
          }
          return true
        }else{
          try {
            let sheetBarWrapperWidth = document.getElementById('sheetBarWrapper').clientWidth
            let sheetBarWidth = document.getElementById('sheetBar').clientWidth
            if(this.data.sheetBar.left*(-1)+sheetBarWrapperWidth>=sheetBarWidth){
              return false
            }
            return true
          }catch (e) {
            return false
          }

        }
      },
      sheetBarLeftRightClick(op){
        this.data.sheetBar.showAnimation = true
        let sheetBarWrapperWidth = document.getElementById('sheetBarWrapper').clientWidth
        let sheetBarWidth = document.getElementById('sheetBar').clientWidth
        let re = document.getElementById('sheetBar').getElementsByClassName('fileBar-fileItem')
        let currentLeft = this.data.sheetBar.left*-1
        if(op){
          let t = 0
          let tn = 0
          for(let i=0;i<re.length;i++){
            if(t+re[i].clientWidth>currentLeft){
              break
            }
            tn+=1
            t+=re[i].clientWidth
          }
          let md = currentLeft - t
          for(let i=tn-1;i>=0;i--){
            if(md+re[i].clientWidth>sheetBarWrapperWidth){
              break
            }
            md+=re[i].clientWidth
          }
          this.data.sheetBar.left+=md
        }else{
          currentLeft += sheetBarWrapperWidth
          let t = 0
          let tn = 0
          for(let i=0;i<re.length;i++){
            if(t+re[i].clientWidth>currentLeft){
              break
            }
            tn+=1
            t+=re[i].clientWidth
          }
          let md = 0
          for(let i = tn;i<re.length;i++){
            if(md+re[i].clientWidth>sheetBarWrapperWidth){
              break
            }
            md+=re[i].clientWidth
          }
          this.data.sheetBar.left -= md
          if(this.data.sheetBar.left*-1+sheetBarWrapperWidth>sheetBarWidth){
            this.data.sheetBar.left = sheetBarWidth*-1 + sheetBarWrapperWidth
          }
        }
        setTimeout(()=>{
          this.data.sheetBar.showAnimation = false
        },500)
      },
      getFileNameFormFileWithoutExt(filePath){
        filePath = tool.getFileNameFromFilePath(filePath)
        filePath = filePath.split('.')
        let re = ''
        for(let i=0;i<filePath.length-(filePath[filePath.length-1]=='JSON'?1:0);i++){
          if(i){
            re+='.'
          }
          re+=filePath[i]
        }
        return re
      },
      async fileList_rename(){
        for(let i=0;i<this.data.files.length;i++){
          if(tool.normalizeBackslashes(this.data.files[i].system.path).indexOf(tool.normalizeBackslashes(this.data.rightMenu.data))!=-1){
            let t = this.data.rightMenu.data.split('.')
            if(t[t.length-1]=='JSON'){
              this.setWarningFormTitleAndContentAndShowForm('无法重命名','此表格已打开。',false,()=>{})
            }else{
              this.setWarningFormTitleAndContentAndShowForm('无法重命名','此文件夹中的表格文件已打开。',false,()=>{})
            }
            return
          }
        }
        if(this.data.rightMenu.data.split('.')[this.data.rightMenu.data.split('.').length-1]=='JSON'){
          if(this.data.new.newFileName==''){
            this.setWarningFormTitleAndContentAndShowForm('无法重命名','请输入新文件名。',false,()=>{})
            return
          }
          if(tool.isValidFileName(this.data.new.newFileName)==false){
            this.setWarningFormTitleAndContentAndShowForm('无法重命名','新文件名不合法。',false,()=>{})
            return
          }
          for(let i=0;i<this.data.fileList.current.length;i++){
            if(this.data.fileList.current[i].children!=undefined){
              continue
            }
            if(this.data.fileList.current[i].name==this.data.new.newFileName + '.JSON'){
              this.setWarningFormTitleAndContentAndShowForm('无法重命名','新文件名已存在。',false,()=>{})
              return
            }
          }
          this.showLoading()
          let re = JSON.parse(await tool.readFile(this.data.rightMenu.data))
          re.name = this.data.new.newFileName
          await tool.updateFileContent(this.data.rightMenu.data,JSON.stringify(re))
          await tool.renameFile(this.data.rightMenu.data,this.data.new.newFileName + '.JSON')
        }else{
          if(this.data.new.newFileName==''){
            this.setWarningFormTitleAndContentAndShowForm('无法重命名','请输入新文件夹名。',false,()=>{})
            return
          }
          if(tool.isValidFolderName(this.data.new.newFileName)==false || this.data.new.newFileName.split('.')[this.data.new.newFileName.split('.').length-1]=='JSON'){
            this.setWarningFormTitleAndContentAndShowForm('无法重命名','新文件夹名不合法。',false,()=>{})
            return
          }
          for(let i=0;i<this.data.fileList.current.length;i++){
            if(this.data.fileList.current[i].children==undefined){
              continue
            }
            if(this.data.fileList.current[i].name==this.data.new.newFileName){
              this.setWarningFormTitleAndContentAndShowForm('无法重命名','新文件夹名已存在。',false,()=>{})
              return
            }
          }
          this.showLoading()
          await tool.renameFolder(this.data.rightMenu.data,this.data.new.newFileName)
        }
        await this.refreshFileList()
        this.hideLoading()
        this.controlForm('rename',0)
      },
      fileBarLetfRightVisible(op){
        if(document.getElementById('fileBarWrapper')==null){
          return false
        }
        if(op){
          if(this.data.fileBar.left>=0){
            return false
          }
          return true
        }else{
          let fileBarWrapperWidth = document.getElementById('fileBarWrapper').clientWidth
          let fileBarWidth = document.getElementById('fileBar').clientWidth
          if(this.data.fileBar.left*(-1)+fileBarWrapperWidth>=fileBarWidth){
            return false
          }
          return true
        }
      },
      fileBarLeftRightClick(op){
        this.data.fileBar.showAnimation = true
        let fileBarWrapperWidth = document.getElementById('fileBarWrapper').clientWidth
        let fileBarWidth = document.getElementById('fileBar').clientWidth
        let re = document.getElementById('fileBar').getElementsByClassName('fileBar-fileItem')
        let currentLeft = this.data.fileBar.left*-1
        if(op){
          let t = 0
          let tn = 0
          for(let i=0;i<re.length;i++){
            if(t+re[i].clientWidth>currentLeft){
              break
            }
            tn+=1
            t+=re[i].clientWidth
          }
          let md = currentLeft - t
          for(let i=tn-1;i>=0;i--){
            if(md+re[i].clientWidth>fileBarWrapperWidth){
              break
            }
            md+=re[i].clientWidth
          }
          this.data.fileBar.left+=md
        }else{
          currentLeft += fileBarWrapperWidth
          let t = 0
          let tn = 0
          for(let i=0;i<re.length;i++){
            if(t+re[i].clientWidth>currentLeft){
              break
            }
            tn+=1
            t+=re[i].clientWidth
          }
          let md = 0
          for(let i = tn;i<re.length;i++){
            if(md+re[i].clientWidth>fileBarWrapperWidth){
              break
            }
            md+=re[i].clientWidth
          }
          this.data.fileBar.left -= md
          if(this.data.fileBar.left*-1+fileBarWrapperWidth>fileBarWidth){
            this.data.fileBar.left = fileBarWidth*-1 + fileBarWrapperWidth
          }
        }
        setTimeout(()=>{
          this.data.fileBar.showAnimation = false
        },500)
      },
      moveArrElemItem(arr,index,op){
        if(op){
          index-=1
          if(index==-1){
            this.setWarningFormTitleAndContentAndShowForm('无法左移','此表格已在最左侧。',false,()=>{})
            return arr
          }
          index+=1
          let t = arr[index-1]
          arr[index-1] = arr[index]
          arr[index] = t
        }else{
          index+=1
          if(index==arr.length){
            this.setWarningFormTitleAndContentAndShowForm('无法右移','此表格已在最右侧。',false,()=>{})
            return arr
          }
          index-=1
          let t = arr[index+1]
          arr[index+1] = arr[index]
          arr[index] = t
        }
        return arr
      },
      sheetBarItemClick(name){
        for(let i=0;i<this.data.files[this.data.currentFile.index-1].data.sheets.length;i++){
          if(this.data.files[this.data.currentFile.index-1].data.sheets[i].name==name){
            this.data.files[this.data.currentFile.index-1].system.currentSheetIndex = i+1
            this.data.files[this.data.currentFile.index-1].system.currentSheetName = name
            return
          }
        }
      },
      async sheet_newSheet(){

        function isValidWorksheetName(name) {
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
        }

        let newSheetName = ''
        if(this.data.new.newFileName==''){
          this.setWarningFormTitleAndContentAndShowForm('无法新建表格','表格名称不能为空。',false,()=>{})
          return
        }else{
          newSheetName = this.data.new.newFileName
          if(isValidWorksheetName(newSheetName)==false){
            this.setWarningFormTitleAndContentAndShowForm('无法新建表格','工作表名称不合法。',true,()=>{

            })
            return
          }
          for(let i=0;i<this.data.files[this.data.currentFile.index-1].data.sheets.length;i++){
            if(this.data.files[this.data.currentFile.index-1].data.sheets[i].name==newSheetName){
              this.setWarningFormTitleAndContentAndShowForm('无法新建表格','工作表已存在。',true,()=>{

              })
              return
            }
          }
        }
        this.controlForm('newSheet',0)

        this.data.files[this.data.currentFile.index-1].data.sheets.push({
          name:newSheetName,
          title:'',
          config:{

          },
          columns:{
            'KpAF0':{
              name:'序号',
              type:'number',
              sumType:'none',
            }
          },
          data:[

          ]
        })
        this.data.files[this.data.currentFile.index-1].system.currentSheetIndex = this.data.files[this.data.currentFile.index-1].data.sheets.length
        this.data.files[this.data.currentFile.index-1].system.currentSheetName = newSheetName
        this.setCurrentFileUnsave()
      },
      setCurrentFileUnsave(){
        this.data.files[this.data.currentFile.index-1].system.unsave = true
      },
      showLoading(){
        this.controlForm('loading',1)
      },
      hideLoading(){
        this.controlForm('loading',0)
      },
      checkIsFileSaved(path){
        for(let i=0;i<this.data.files.length;i++){
          if(this.data.files[i].system.path==path){
            return this.data.files[i].system.unsave==false
          }
        }
      },
      async file_closeSingle(){
        if(this.checkIsFileSaved(this.data.rightMenu.data)==false){
          this.setWarningFormTitleAndContentAndShowForm('提示','此表格文件未保存，是否保存并关闭？',true,async ()=>{
            this.showLoading()
            await this.file_saveSingle()
            await this.file_closeFile(this.data.rightMenu.data)
            this.hideLoading()
          },true,{
            cancel: '取消',
            ok:'保存并关闭',
            secondary: '不保存并关闭'
          },async ()=>{
            this.showLoading()
            await this.file_closeFile(this.data.rightMenu.data)
            this.hideLoading()
          })
          return
        }
        this.controlForm('loading',true)
        await this.file_closeFile(this.data.rightMenu.data)
        this.controlForm('loading',false)
      },
      async file_saveSingle(){
        this.controlForm('loading',true)
        for(let i=0;i<this.data.files.length;i++){
          if(this.data.files[i].system.path==this.data.rightMenu.data){
            if(this.data.files[i].system.unsave){
              await tool.updateFileContent(this.data.rightMenu.data,JSON.stringify(this.data.files[i].data))
              this.data.files[i].system.unsave = false
            }
            break
          }
        }
        this.controlForm('loading',false)
      },
      async file_closeAll(){
          for(let i = this.data.files.length-1;i>=0;i--){
            if(this.data.files[i].system.unsave){
              this.setWarningFormTitleAndContentAndShowForm('提示','有表格文件未保存，是否将所有文件保存并全部关闭？',true,async ()=>{
                this.showLoading()
                await this.file_saveAll()
                await this.file_closeAll()
                this.hideLoading()
              },true,{
                cancel: '取消',
                ok:'全部保存并关闭',
                secondary: '全部不保存并关闭'
              },async ()=>{
                this.showLoading()
                for(let i1=0;i1<this.data.files.length;i1++){
                  this.data.files[i1].system.unsave = false
                }
                await this.file_closeAll()
                this.hideLoading()
              })
              return false
            }
        }
        this.controlForm('loading',true)
        for(let i = this.data.files.length-1;i>=0;i--){
          this.file_closeFile(this.data.files[i].system.path)
        }
        this.controlForm('loading',false)
        return true
      },
      async file_saveAllClose(){
        this.showLoading()
        await this.file_saveAll()
        await this.file_closeAll()
        this.hideLoading()
      },
      async file_saveAll(){
        this.controlForm('loading',true)
        for(let i=0;i<this.data.files.length;i++){
          if(this.data.files[i].system.unsave==true){
            await tool.updateFileContent(this.data.files[i].system.path,JSON.stringify(this.data.files[i].data))
          }
        }
        for(let i=0;i<this.data.files.length;i++){
          this.data.files[i].system.unsave = false
        }
        this.controlForm('loading',false)
      },
      async fileList_moveClick(){
        this.controlForm('loading',true)
        await tool.moveFile(this.data.rightMenu.data,this.data.new.currentFolderPath + '\\' + tool.getFileNameFromFilePath(this.data.rightMenu.data))
        await this.refreshFileList()
        this.controlForm('loading',false)
        this.controlForm('move',false)
      },
      async fileList_move(){
        for(let i=0;i<this.data.files.length;i++){
          if(tool.normalizeBackslashes(this.data.files[i].system.path).indexOf(tool.normalizeBackslashes(this.data.rightMenu.data))!=-1){
            let t = this.data.rightMenu.data.split('.')
            if(t[t.length-1]=='JSON'){
              this.setWarningFormTitleAndContentAndShowForm('无法移动','此表格已打开。',false,()=>{})
            }else{
              this.setWarningFormTitleAndContentAndShowForm('无法移动','此文件夹中的表格文件已打开。',false,()=>{})
            }
            return
          }
        }
        this.data.new.showSelectFolder = false
        setTimeout(()=>{
          this.data.new.showSelectFolder = true
          this.controlForm('move',1)
        },1)

      },
      async fileList_delete(){
        for(let i=0;i<this.data.files.length;i++){
          if(tool.normalizeBackslashes(this.data.files[i].system.path).indexOf(tool.normalizeBackslashes(this.data.rightMenu.data))!=-1){
            let t = this.data.rightMenu.data.split('.')
            if(t[t.length-1]=='JSON'){
              this.setWarningFormTitleAndContentAndShowForm('无法删除','此表格已打开。',false,()=>{})
            }else{
              this.setWarningFormTitleAndContentAndShowForm('无法删除','此文件夹中的表格文件已打开。',false,()=>{})
            }
            return
          }
        }
        let t = this.data.rightMenu.data.split('.')
        if(t[t.length-1]=='JSON'){
          this.setWarningFormTitleAndContentAndShowForm('删除','删除后不可恢复，是否继续？',true,async ()=>{
            this.controlForm('loading',true)
            await tool.deleteFile(this.data.rightMenu.data)
            await this.refreshFileList()
            this.controlForm('loading',false)
          })
        }else{
          this.setWarningFormTitleAndContentAndShowForm('删除','删除后不可恢复，是否继续？',true,async ()=>{
            this.controlForm('loading',true)
            await tool.deleteFolder(this.data.rightMenu.data)
            await this.refreshFileList()
            this.controlForm('loading',false)
          })
        }
      },
      checkMouseDownEvent(event){
        if(event.button==2){
          this.data.rightMenu.allowShow = true
        }else{
          this.data.rightMenu.allowShow = false
        }
      },
      hideRightMenu(){
        this.controlForm('rightMenu',0)
        this.data.rightMenu.show = false
      },
      showRightMenu(rightMenuName,data,positionModifyX,positionModifyY){
        if(this.data.rightMenu.allowShow==false)return
        this.data.rightMenu.top = this.data.rightMenu.topT
        this.data.rightMenu.left = this.data.rightMenu.leftT
        this.data.rightMenu.name = rightMenuName
        this.data.rightMenu.data = data
        this.controlForm('rightMenu',1)
        this.data.rightMenu.show = true
        if(positionModifyX!=undefined){
          this.data.rightMenu.left += positionModifyX
        }
        if(positionModifyY!=undefined){
          this.data.rightMenu.top += positionModifyY
        }
      },
      getMousePosition(event){
        this.data.rightMenu.leftT = event.clientX
        this.data.rightMenu.topT = event.clientY
        this.data.window.width = document.body.clientWidth
        this.data.window.height = document.body.clientHeight
        if(this.data.rightMenu.show){
          return
        }
        this.data.rightMenu.left = event.clientX
        this.data.rightMenu.top = event.clientY
      },
      async file_closeFile(path){
        for(let i=0;i<this.data.files.length;i++){
          if(this.data.files[i].system.path==path){
            if((i+1)<this.data.currentFile.index){
              this.data.currentFile.index-=1
            }
            this.data.files = tool.deleteArrElemByIndex(this.data.files,i)
            break
          }
        }

        if(this.data.files.length==0){
          switch (this.data.setting.readData.closeAllFiles) {
            case 'exit':
              remote.app.exit()
              break
            case 'showStart':
              this.showStartPage()
              break
          }
        }else{
          if(path==this.data.currentFile.path){
            this.data.currentFile.path = this.data.files[0].system.path
            this.data.currentFile.index = 1
          }
        }
      },
      fileBarTitleGetPath(filePath){
        filePath = filePath.split('.')
        let re = ''
        for(let i=0;i<filePath.length-1;i++){
          re+=filePath[i]
        }
        return re
      },
      file_fileBarItemClick(path){
        for(let i=0;i<this.data.files.length;i++){
          if(this.data.files[i].system.path==path){
            this.data.currentFile.path = path
            this.data.currentFile.index = i+1
            break
          }
        }
      },
      async file_open(filePath){
        filePath = tool.normalizeBackslashes(filePath)
        function getFileNameFromPath(filePath) {  
          const lastSlashIndex = filePath.lastIndexOf('/');  
          const lastBackslashIndex = filePath.lastIndexOf('\\');  
          const separatorIndex = Math.max(lastSlashIndex, lastBackslashIndex);  
          if (separatorIndex === -1) {  
            return filePath;  
          }

          return filePath.slice(separatorIndex + 1);  
        }
        let fileName = this.getFileNameWithoutExtension(getFileNameFromPath(filePath))
        let re = {
          data:JSON.parse(await tool.readFile(filePath)),
          system:{
            path:filePath,
            unsave:false,
            currentSheetIndex:(JSON.parse(await tool.readFile(filePath)).sheets.length==0?0:1),
            currentSheetName:(JSON.parse(await tool.readFile(filePath)).sheets.length==0?'':JSON.parse(await tool.readFile(filePath)).sheets[0].name),
            recentEditTime:tool.getFormattedLastModifiedTime(filePath),
            operationStack:[],
            operationStackIndex:0,
            search:{
              show:false,
              text:'',
              key:'',
              selectKeys:[],
              results:[],
              showResult:false,
            },
            aiAides:{
              show:false,
              contents:[

              ]
            },
          }
        }
        let opened = false
        for(let i=0;i<this.data.files.length;i++){
          if(this.data.files[i].system.path==filePath){
            opened = true
            this.data.currentFile.index = i+1
            this.data.currentFile.path = filePath
            break
          }
        }
        if(opened){

        }else{
          this.data.files.push(re)
          this.data.currentFile.index = this.data.files.length
          this.data.currentFile.path = filePath
        }
        this.hideStartPage()
      },
      async newFolderClick(){
        this.data.new.newFolderName = ''
        this.controlForm('newFolder',true)
        this.data.new.currentFolderPath = this.data.setting.readData.saveDrive + '\\ExcelEasyData' + this.getCurrentFileListFolderPath()
      },
      getCurrentFileListFolderPath(){
        let re = ''
        for(let i=0;i<this.data.fileList.path.length;i++){
          re+='\\'
          re+=this.data.fileList.path[i].name
        }
        return re
      },
      fileListNavItemClick(item){
        let t = this.data.fileList.all
        for(let i=0;i<item.index;i++){
          for(let i1=0;i1<t.length;i1++){
            if(t[i1].children==undefined)continue
            if(t[i1].name==this.data.fileList.path[i].name){
              t = t[i1].children
              break
            }
          }
        }
        this.data.fileList.current = t
        for(let i = this.data.fileList.path.length;i>item.index;i--){
          this.data.fileList.path = tool.deleteArrElemByIndex(this.data.fileList.path,i-1)
        }
      },
      fileListBackToStart(){
        this.data.fileList.current = this.data.fileList.all
        this.data.fileList.path = []
      },
      async fileListItemClick(name){
        for(let i=0;i<this.data.fileList.current.length;i++){
          if(this.data.fileList.current[i].name==name){
            this.data.fileList.current = this.data.fileList.current[i].children
            break
          }
        }
        this.data.fileList.path.push({
          name:name,
          index:this.data.fileList.path.length+1
        })
      },
      getFileNameWithoutExtension(fileName) {  
        // 获取文件名的最后一个点的索引  
        const lastDotIndex = fileName.lastIndexOf('.');  

        // 如果没有点（即没有扩展名），则返回整个文件名  
        if (lastDotIndex === -1) {  
          return fileName;  
        }  

        // 否则，返回最后一个点之前的所有内容  
        return fileName.slice(0, lastDotIndex);  
      },
      getFileExtension(fileName) {
        const parts = fileName.split('.')
        if (parts.length === 1) {  
          return '';  
        }  
        return parts.pop();  
      },
      formatDate(date) {  
        const year = date.getFullYear();  
        const month = String(date.getMonth() + 1).padStart(2, '0'); // 月份从0开始，需要加1，并补0  
        const day = String(date.getDate()).padStart(2, '0'); // 日期补0  
        const hours = String(date.getHours()).padStart(2, '0'); // 小时补0  
        const minutes = String(date.getMinutes()).padStart(2, '0'); // 分钟补0  
        const seconds = String(date.getSeconds()).padStart(2, '0'); // 秒补0  

        return `${year}年${month}月${day}日 ${hours}:${minutes}:${seconds}`;  
      },
      async getFileList(){
        let re = await tool.getFolderTree(this.data.setting.readData.saveDrive + "\\ExcelEasyData",true)
        for(let i=re.length-1;i>=0;i--){
          if(re[i].children==undefined){
            if(this.getFileExtension(re[i].name)!='JSON'){
              re = tool.deleteArrElemByIndex(re,i)
            }
          }
        }
        return re
      },
      async addFolder(){
        if(this.data.new.newFolderName==''){
          this.setWarningFormTitleAndContentAndShowForm('无法新建文件夹','请输入文件夹名称。',false,()=>{})
          return
        }
        if(tool.isValidFolderName(this.data.new.newFolderName)==false || this.data.new.newFolderName.split('.')[this.data.new.newFolderName.split('.').length-1]=='JSON'){
          this.setWarningFormTitleAndContentAndShowForm('无法新建文件夹','文件夹名称不合法。',false,()=>{})
          return
        }
        this.controlForm('loading',true)
        let re = await tool.getDirectoriesInFolder(this.data.new.currentFolderPath)
        for(let i=0;i<re.length;i++){
          if(re[i]==this.data.new.newFolderName){
            this.setWarningFormTitleAndContentAndShowForm('无法新建文件夹','当前目录已有名为"' + this.data.new.newFolderName + '"的文件夹。',false,()=>{})
            this.controlForm('loading',false)
            return
          }
        }
        await tool.createFolder(this.data.new.currentFolderPath + '\\' + this.data.new.newFolderName)
        await this.refreshFileList()
        this.controlForm('loading',false)
        this.controlForm('newFolder',false)
        await this.reloadFolderList()
        
      },
      pathChanged(e){
        this.data.new.currentFolderPath = e.path
        this.data.new.currentFolderFileNames = e.fileNames
      },
      async addFile(){
        if(this.data.new.newFileName==''){
          this.setWarningFormTitleAndContentAndShowForm('无法新建表格文件','请输入表格文件名。',false,()=>{})
          return
        }
        if(tool.isValidFileName(this.data.new.newFileName)==false){
          this.setWarningFormTitleAndContentAndShowForm('无法新建表格文件','表格文件名不合法。',false,()=>{})
          return
        }
        for(let i=0;i<this.data.new.currentFolderFileNames.length;i++){
          if(this.data.new.currentFolderFileNames[i]==this.data.new.newFileName + ".JSON"){
            this.setWarningFormTitleAndContentAndShowForm('无法新建表格文件','当前文件夹下已有文件名为' + this.data.new.newFileName + '的表格文件。',false,()=>{})
            return
          }
        }
        this.controlForm('loading',true)
        await tool.createFileWithContent(this.data.new.currentFolderPath + "\\" + this.data.new.newFileName + ".JSON",JSON.stringify({
          name:this.data.new.newFileName,
          sheets:[
            {
              name:'Sheet 1',
              title:'',
              config:{

              },
              columns:{
                'KpAF0':{
                  name:'序号',
                  type:'number',
                  sumType:'none',
                }
              },
              data:[

              ]
            }
          ]
        }))
        await this.refreshFileList()
        if(this.data.showStartPage==2){
          let t = this.data.new.currentFolderPath.split('\\')
          t = tool.deleteArrElemByIndex(t,0)
          t = tool.deleteArrElemByIndex(t,0)
          for(let i=0;i<t.length;i++){
            t[i] = {
              name:t[i]
            }
          }
          this.data.fileList.path = t
        }
        await this.refreshFileList()
        this.file_open(this.data.new.currentFolderPath + "\\" + this.data.new.newFileName + ".JSON")
        this.controlForm('loading',false)
        this.controlForm('new',false)
      },
      reloadFolderList(){
        return new Promise((resolve, reject) => {
          this.data.new.showSelectFolder = false
          setTimeout(() => {
            this.data.new.showSelectFolder = true
            return resolve()
          }, 1);
        })
      },
      async newButtonClick(){
        this.data.new.defaultPath = this.data.setting.readData.saveDrive
        this.data.new.newFileName = ''
        this.data.new.showSelectFolder = false
        this.data.new.currentFolderPath = this.data.setting.readData.saveDrive + '\\ExcelEasyData'
        setTimeout(() => {
          this.data.new.showSelectFolder = true
          this.controlForm('new',true)
        }, 1);
      },
      setWarningFormTitleAndContentAndShowForm(title,content,showCancel,callBack=()=>{},showSecondary=false,buttonTexts = {
        cancel:'取消',
        ok:'确定',

        secondary:'undefined'
      },secondaryCallBack=()=>{}){
        this.data.warning.title = title
        this.data.warning.content = content
        this.data.warning.showCancel = showCancel
        this.data.warning.callBack = callBack
        this.data.warning.buttonTexts = buttonTexts
        this.data.warning.showSecondary = showSecondary
        this.data.warning.secondaryCallBack = secondaryCallBack
        this.controlForm('warning',1)
      },
      getAppPath(){
        const userHomeDir = os.homedir();
        const documentsDir = path.join(userHomeDir, 'Documents');
        return documentsDir
      },
      getSettingFileName(){
        return 'ExcelEasyConfig.JSON'
      },
      async saveSetting(){
        for(let i=0;i<this.data.files.length;i++){
          if(this.data.files[i].system.unsave==true){
            this.hideLoading()
            this.setWarningFormTitleAndContentAndShowForm('无法保存设置','当前有未保存的文件，请全部保存后再试。',false,()=>{})
            return
          }
        }
        this.controlForm('loading',true)
        // let appPath = remote.app.getAppPath()
        let appPath = this.getAppPath() + '\\'
        this.data.setting.readData.saveDrive = this.data.setting.form.saveDrive
        this.data.setting.readData.closeAllFiles = this.data.setting.form.closeAllFiles
        this.data.setting.readData.theme = this.data.setting.form.theme
        this.data.setting.readData.showBackgroundImage = this.data.setting.form.showBackgroundImage
        this.data.setting.readData.gptKey = this.data.setting.form.gptKey
        await tool.updateFileContent(appPath + "\\ExcelEasyConfig.JSON",JSON.stringify(this.data.setting.readData))
        this.data.setting.readData = undefined
        await tool.waitSeconds(1)
        // await this.checkSetting()
        // this.controlForm('loading',false)
        this.controlForm('setting',false)
        // location.reload()
        remote.app.relaunch()
        remote.app.exit()
      },
      async checkSetting(){
        // let appPath = remote.app.getAppPath()
        let appPath = this.getAppPath() + '\\'
        //设置文件
        if(await tool.checkPathExists(appPath + "\\ExcelEasyConfig.JSON")){
          //读取设置
          this.data.setting.readData = JSON.parse(await tool.readFile(appPath + "\\ExcelEasyConfig.JSON"))
          this.data.setting.form.saveDrive = this.data.setting.readData.saveDrive
          this.data.setting.form.closeAllFiles = this.data.setting.readData.closeAllFiles
          this.data.setting.form.theme = (this.data.setting.readData.theme==undefined?'default':this.data.setting.readData.theme)
          this.data.setting.form.showBackgroundImage = (this.data.setting.readData.showBackgroundImage==undefined?false:this.data.setting.readData.showBackgroundImage)
          this.data.setting.form.gptKey = (this.data.setting.readData.gptKey==undefined?'':this.data.setting.readData.gptKey)
          this.data.aiAides.contents = (this.data.setting.readData.aiAidesContents==undefined?[]:this.data.setting.readData.aiAidesContents)
        }else{
          //新建设置
          await tool.createFileWithContent(appPath + "\\ExcelEasyConfig.JSON",JSON.stringify({
            saveDrive:(await this.getSaveDrive()),
            closeAllFiles:'exit',
            theme:'default',
            showBackgroundImage:false,
            gptKey:''
          }))
          await this.checkSetting()

          return
        }
        try {
          //数据存储位置
          if((await tool.checkPathExists(this.data.setting.readData.saveDrive + "\\ExcelEasyData"))==false){
            await tool.createFolder(this.data.setting.readData.saveDrive + "\\ExcelEasyData")
          }
        }catch (e){

        }

      },
      async getSaveDrive(){
        let re = await tool.getAllDrive()
        return re[0]
      },
      async prepareSettingData(){
        //获取所有驱动器
        this.data.setting.drives = await tool.getAllDrive()
        this.data.setting.form.saveDrive = this.data.setting.readData.saveDrive

      },
      controlForm(id,op){
        tool.controlForm(id,op)
      },
      async controlWindow(op){
        const win = remote.getCurrentWindow();
        switch (op) {
          case 'max':
            if(win.isMaximized()){
              win.restore()
              this.data.cwm = false
            }else{
              win.maximize()
              this.data.cwm = true
            }
            break;
          case 'min':
            win.minimize()
            break;
          case 'close':
            for(let i = this.data.files.length-1;i>=0;i--){
              if(this.data.files[i].system.unsave){
                this.setWarningFormTitleAndContentAndShowForm('提示','有表格文件未保存，是否将所有文件保存并全部关闭？',true,async ()=>{
                  this.showLoading()
                  await this.file_saveAll()
                  await this.file_closeAll()
                  this.hideLoading()
                  remote.app.exit()
                },true,{
                  cancel: '取消',
                  ok:'全部保存并关闭',
                  secondary: '全部不保存并关闭'
                },async ()=>{
                  this.showLoading()
                  for(let i1=0;i1<this.data.files.length;i1++){
                    this.data.files[i1].system.unsave = false
                  }
                  await this.file_closeAll()
                  this.hideLoading()
                  remote.app.exit()
                })
                return false
              }
            }
            remote.app.exit()
            break;
        }
      },
      getGreetingBasedOnTime() {  
          const date = new Date();  
          const hour = date.getHours();  

          if (hour >= 0 && hour < 12) {  
              return "早上好";  
          } else if (hour >= 12 && hour < 14) {  
              return "中午好";  
          } else if (hour >= 14 && hour < 18) {  
              return "下午好";  
          } else if (hour >= 18 && hour < 20) {  
              return "傍晚好";  
          } else {  
              return "晚上好";  
          }  
      },
      async hideStartPage(){
        this.data.showStartPage = 1
        setTimeout(() => {
          this.data.showStartPage = 0
        }, 200);
      },
      async showStartPage(){
        this.data.showStartPage = 1
        setTimeout(() => {
          this.data.showStartPage = 2
        }, 1);
      },
      async refreshFileList(){
        let re = await this.getFileList()
        this.data.fileList.all = re
        this.data.fileList.current = re
        let patht = this.data.fileList.path
        this.data.fileList.path = []
        for(let i=0;i<patht.length;i++){
          await this.fileListItemClick(patht[i].name)
        }
      }
    },

    async mounted() {

      // remote.getCurrentWebContents().openDevTools()
      //准备程序
      try {
        this.controlForm('loading',true)
        await this.checkSetting()
        await this.prepareSettingData()
        await this.refreshFileList()
        this.controlForm('loading',false)
        this.data.appStarted = true
      }catch(e){
        this.controlForm('loading',false)
        this.setWarningFormTitleAndContentAndShowForm('请插入可移动设备','未找到你设置的存储位置的驱动器，请插入此驱动器，然后重启软件，或者修改存储位置。',false,()=>{
          remote.app.relaunch()
          remote.app.exit()
          return
        },true,{
          cancel: '取消',
          ok:'重启软件',
          secondary: '修改存储位置'
        },()=>{
          this.controlForm('setting',true)
        })
        return
      }

      if(this.data.setting.readData.showBackgroundImage){
        try {
          this.getBackground()
        }catch (e) {


        }

      }


      setInterval(() => {
        if(remote.getCurrentWindow().isMaximized()){
          this.data.cwm = true
        }else{
          this.data.cwm = false
        }

        try {
          let fileBarWrapperWidth = document.getElementById('fileBarWrapper').clientWidth
          let fileBarWidth = document.getElementById('fileBar').clientWidth
          let fileBarLeft = this.data.fileBar.left*-1
          if(fileBarLeft + fileBarWrapperWidth>fileBarWidth){
            if(fileBarWrapperWidth>fileBarWidth){
              this.data.fileBar.left = 0
            }else{
              this.data.fileBar.left = fileBarWrapperWidth - fileBarWidth
            }
          }

          let sheetBarWrapperWidth = document.getElementById('sheetBarWrapper').clientWidth
          let sheetBarWidth = document.getElementById('sheetBar').clientWidth
          let sheetBarLeft = this.data.sheetBar.left*-1
          if(sheetBarLeft + sheetBarWrapperWidth>sheetBarWidth){
            if(sheetBarWrapperWidth>sheetBarWidth){
              this.data.sheetBar.left = 0
            }else{
              this.data.sheetBar.left = sheetBarWrapperWidth - sheetBarWidth
            }
          }

          this.data.tableWrapperScrollLeft = document.getElementById('tableScrollPage').scrollLeft
        }catch (e) {

        }

        this.data.tableOnScrollTopBar.show = this.showTableOnScrollTop()
      }, 10);
    },
  }
</script>

<style>

.rightMenuLabel{
  font-size: 12px;
  padding-left: 5px;
  padding-top: 5px;
  padding-bottom: 5px;
  color: gray;
  padding-right: 5px;
}

.rightMenuDivider{
  width: 100%;
  height: 0.2px;
  background-color: lightgray;
}

.table-line:hover{
  background-color: rgb(242,242,242);
}

.table-line-selected{
  background-color: rgba(0,120,215,0.2);
}

td{
  text-align: center;
  font-size: 14px;
  padding: 10px 10px;
}

.mcell{
  width: 100%;
  border-radius: 5px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.2);
  margin-bottom: 20px;
}

.mcell:hover{
  background-color: rgba(128,128,128,0.1);
}

.mcell:active{
  background-color: rgba(128,128,128,0.05);
}

.column{
  min-width: 60px;
  display: inline-block;
  text-align: center;
  background-color: red;
  height: calc(100vh - 250px);
  padding-bottom: 80px;
  position: relative;
}

.cell-columnTitle{
  font-size: 18px;
  font-family: 黑体;
}

.cell-title-black{
  color: lightgray;
}

.cell{
  cursor: cell;
  border-right: 1px solid black;
  border-bottom: 1px solid black;
  padding-top: 5px;
  padding-bottom: 5px;
}

.opBarItem svg{
  margin-top: 6px;
  margin-right: 5px;
  display: inline-block;
}

.opBarItem span{
  position: relative;
  top: -3px;
}

.opBarItem{
  height: 28px;
  margin-top: 4px;
  border-radius: 5px;
  line-height: 28px;
  font-size: 14px;
  padding: 0 5px;
  margin-left: 2px;
  margin-right: 2px;
  display: inline-block;
}

.opBarItem-disabled{
  color: gray;
  background-color: white !important;
}

.opBarItem:hover{
  background-color: rgba(0,0,0,0.1);
}

.opBarItem:active{
  background-color: rgba(0,0,0,0.05);
}

.rightMenuButton{
  padding: 10px 10px;
  font-size: 14px;
}

.rightMenuButton:hover{
  background-color: rgb(233, 233, 233);
}

.rightMenuButton:active{
  background-color: #ededed;
}

.rightMenu{
  border-radius: 5px;
  box-shadow: 0 4px 10px rgba(0,0,0,0.3);
  display: inline-block;
  min-width: 100px;
  font-size: 14px;
  background-color: white;
  overflow: hidden;
}

.fileBar-fileItem{
  line-height: 36px;
  padding: 0 10px;
  font-size: 14px;
  min-width: 100px;
  padding-right: 30px;
  position: relative;
  background-color: white;
}

.fileBar-fileItem-selected{
  border-bottom: 2px solid rgb(0,120,215);
  height: 36px;
  background-color: rgb(242,242,242);
}

.fileBar-fileItem:hover{
  background-color: rgb(230,230,230);
}

.fileBar-fileItem:active{
  background-color: rgb(242,242,242);
}

.fileItemTitle{
  
}

.fileItemTitle:hover{
  color:rgb(0,120,215)
}

.form-footer{
  width: 100%;
  padding: 20px 20px;
  padding-top: 10px;
  text-align: right;
}

.form-content{
  width: 100%;
  padding: 20px 20px;
  padding-top: 10px;
  height: calc(100% - 56.67px - 71.33px);
  overflow-y: auto;
  font-size: 14px;
}

.form-title{
  padding: 20px 20px;
  padding-bottom: 10px;
  font-size: 20px;
  font-weight: bold;
}

.form-window{
  width: 90vw;
  max-width: 500px;
  height: 600px;
  max-height: 80vh;
  border-radius: 10px;
  background-color: white;
  box-shadow: 0 4px 10px rgba(0,0,0,0.1);
  transform: translate(-50%,-50%);
  position: relative;
  overflow: hidden;
}

.form-bg{
  width: 100vw;height: 100vh;position: fixed;left: 0;top: 0;background-color: rgba(0,0,0,0.1);z-index: 40000000 !important;padding-left: 50vw;padding-top: 50vh;
}

.min{
  background-image: url('./assets/min.png');
}

.max{
  background-image: url('./assets/max.png');
}

.restore{
  background-image: url('./assets/restore.png');
}

.close{
  background-image: url('./assets/close.png') !important;
}

.min_dark{
  background-image: url('./assets/min_dark.png');
}

.max_dark{
  background-image: url('./assets/max_dark.png');
}

.restore_dark{
  background-image: url('./assets/restore_dark.png');
}

.close_dark{
  background-image: url('./assets/close_dark.png') !important;
}

.close:hover{
  background-image: url('./assets/close_dark.png') !important;
}

.close:active{
  background-image: url('./assets/close_dark.png') !important;
}

.close_dark:hover{
  background-image: url('./assets/close_dark.png') !important;
}

.close_dark:active{
  background-image: url('./assets/close_dark.png') !important;
}

.button{
  
}

.button:hover{
  background-color: rgba(0,0,0,0.1);
}

.button:active{
  background-color: rgba(0,0,0,0.05);
}

.button_dark:hover{
  background-color: rgba(255,255,255,0.1);
}

.button_dark:active{
  background-color: rgba(255,255,255,0.05);
}

.close-button:hover{
  background-color: rgb(196, 43, 28);
}

.close-button:active{
  background-color: #c83c31;
}

.white-background{
  background-image: url("./assets/trans.png") !important;
}

</style>