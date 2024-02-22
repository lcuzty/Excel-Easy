<template>
  <div style="width: 100%;height: 350px;overflow-y:auto;overflow-x:hidden">
    <div class="button" style="width:100%;padding:10px 10px;border-radius:3px" v-if="data.path.length!=0" @click="changeToLast();pathChanged();">
        {{ '返回 ' + getPathString() }}
    </div>
    <div v-for="item in data.folders">

        <div class="button" style="width:calc(100% - 0px);padding:10px 10px;display:inline-block;border-radius:3px"
            @click="async ()=>{
                await func.waitSeconds(0.1)
                data.path.push(item.name)
                data.folders = item.children
                pathChanged()
            }"
        >
            {{ item.name }}
        </div>

    </div>
    <p v-if="data.folders.length==0" style="margin-top:5px">此文件夹没有子文件夹。</p>
    
  </div>
</template>

<script>
import tool from '@/utlis/tool'

export default {
  components:{
    
  },
  data() {
    return {
        data:{
            folders:[],
            path:[],
            foldersAll:[]
        },
        func:{
            waitSeconds:undefined
        }
    };
  },
  methods: {
    async pathChanged(){
        this.$emit('pathChanged',{
            path: this.getRePathString(),
            fileNames: await this.getReFileNames()
        })
    },
    async getReFileNames(){
        return await tool.getFilesInFolder(this.getRePathString())
    },
    getRePathString(){
        let re = this.parentFolderPath + '\\ExcelEasyData'
        for(let i=0;i<this.data.path.length;i++){
            re+='\\'
            re+=this.data.path[i]
        }
        return re
    },
    getPathString(){
        let re = this.parentFolderPath.slice(0,this.parentFolderPath.length-1) + "盘\\"
        for(let i=0;i<this.data.path.length-1;i++){
            re+='\\'
            re+=this.data.path[i]
        }
        return re
    },
    changeToLast(){
        let t = JSON.parse(JSON.stringify(this.data.foldersAll))
        for(let i=0;i<this.data.path.length-1;i++){
            for(let i1=0;i1<t.length;i1++){
                if(t[i1].name==this.data.path[i]){
                    t = t[i1].children
                    break
                }
            }
        }
        this.data.folders = t
        this.data.path = tool.deleteArrElemByIndex(this.data.path,this.data.path.length-1)
    },
    async getTree(){
        this.data.folders = await tool.getFolderTree(this.parentFolderPath + '\\ExcelEasyData')
        this.data.foldersAll = JSON.parse(JSON.stringify(this.data.folders))
    },
  },
  async mounted() {
    this.func.waitSeconds = tool.waitSeconds
    await this.getTree()
    this.pathChanged()
  },
  props: {
    parentFolderPath:''
  }
};
</script>

<style>

</style>