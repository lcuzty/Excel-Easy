<template>
  <div v-if="show" style="width: 100%;padding-bottom: 10px;">
    <div v-if="config.position==1" style="text-align: right;width: 100%">
      <div style="font-size: 15px;padding: 10px 10px;max-width: 80%;border-radius: 10px;background-color: white;box-shadow: 0 1px 3px rgba(0,0,0,0.2);user-select: text;text-align: left;color: white;background: linear-gradient(135deg,#2b67e2,#2447db);display: inline-block">
        {{ text }}
      </div>
    </div>
    <div v-if="config.position==-1" style="font-size: 15px;padding: 10px 10px;max-width: 80%;border-radius: 10px;background-color: white;box-shadow: 0 1px 3px rgba(0,0,0,0.2);user-select: text;display: inline-block">
      {{ text }}
    </div>
  </div>
</template>

<script>
import tool from '../utlis/tool'
export default {
  data() {
    return {
      text:'',
      show:false,
    };
  },
  methods: {

  },
  async mounted() {
    if(this.config==undefined){
      return
    }
    this.show = true
    if(this.config.showAnimation){
      this.$emit('disableAnimation',undefined)
      for(let i=0;i<this.config.text.length;i++){
        this.text += this.config.text.slice(i,i+1)
        try {
          document.getElementById('aichat').scrollTop = 10000000000
        }catch(e){

        }
        await tool.waitSeconds(0.03)
      }
      this.$emit('endLoading',undefined)
    }else{
      this.text = this.config.text
    }
  },
  props: {
    config:undefined,
  }
};
</script>

<style>

</style>