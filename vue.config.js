const { defineConfig } = require('@vue/cli-service')
module.exports = defineConfig({
  transpileDependencies: true,
  pluginOptions:{
    electronBuilder:{
      nodeIntegration:true,
      builderOptions: {
        productName: "ExcelEasy",
        directories: {
          output: "dist_electron"
        },
        nsis: { oneClick: false, 
                allowToChangeInstallationDirectory: true,
                shortcutName: "Excel Easy", 
                createStartMenuShortcut:true,
                runAfterFinish:false,
              },
        // win: { target: "nsis", icon: "public/dmslogo.png" } 
      }
    }
    // electronBuilder: {
    //   builderOptions: {
    //     'productName': 'all electron',//生成exe的名字
    //     "appId": "com.xi.www",//包名  
    //     "copyright": "xi",//版权信息
    //     "directories": { // 输出文件夹
    //       "output": "electron_output",
    //     },
    //     "nsis": {
    //       "oneClick": false, // 是否一键安装
    //       "allowElevation": true, // 允许请求提升。若为false，则用户必须使用提升的权限重新启动安装程序。
    //       "allowToChangeInstallationDirectory": true, //是否允许修改安装目录
    //       "installerIcon": "./build/icons/icon.ico",// 安装时图标
    //       "uninstallerIcon": "./build/icons/icon.ico",//卸载时图标
    //       "installerHeaderIcon": "./build/icons/icon.ico", // 安装时头部图标
    //       "createDesktopShortcut": true, // 是否创建桌面图标
    //       "createStartMenuShortcut": true,// 是否创建开始菜单图标
    //       "shortcutName": "all-electron", // 快捷方式名称
    //       "runAfterFinish": false,//是否安装完成后运行
    //     },
    //     "win": {
    //       "icon": "build/icons/icon.ico",//图标路径
    //       "target": [
    //         {
    //           "target": "nsis", //利用nsis制作安装程序
    //           "arch": [
    //             "x64", //64位
    //             // "ia32" //32位
    //           ]
    //         }
    //       ]
    //     }
    //   }
    // }
  },

    devServer: {
    proxy: {
        '/chat': {
            target: 'http://8.130.76.76:5000',
                changeOrigin: true
        }
    }
}

})
