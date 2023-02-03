<template>
  <div>
    <el-container>
      <el-header>
        <el-row>
          <el-col :span="24">
            <div style="line-height: 60px">
              <i class="el-icon-data-analysis"></i> 数据处理中心
            </div>
          </el-col>
        </el-row>
      </el-header>
      <el-main>
        <!--
          <nuxt-link to="/upload">另个一上传</nuxt-link>
          <nuxt-link to="/upload/123">123</nuxt-link>
        -->
        <el-row>
          <el-col :span="12">
            <el-upload class="upload-demo" ref="upload" action="http://127.0.0.1:8080/posts/" accept=".xlsx,xls"
              :limit="1" :on-preview="handlePreview" :on-remove="handleRemove" :file-list="fileList"
              :auto-upload="false" :on-change="handle">
              <el-button slot="trigger" size="small" type="primary">选取文件</el-button>
              <el-button size="small" type="success" @click="exportWord">本地解析</el-button>
              <el-button style="margin-left: 10px" size="small" type="danger" @click="submitUpload">服务器解析</el-button>
              <!-- <el-button size="small" type="warning" @click="showWord">预览模板</el-button> -->
              <div slot="tip" class="el-upload__tip">只能上传xlsx/xls文件</div>
            </el-upload>
          </el-col>
           <el-col :span="6">
              <input type="file" name="file" @change="changeFile"   />
              <!-- <input type="file" name="file" @change="changeFiles" webkitdirectory  /> -->
              <!-- <el-input v-model="filesPath" placeholder="请输入模板存放路径，例如:C:\Users\xxx\Desktop\model"></el-input> -->
            </el-col>
            <!-- <el-col :span="2">
              <el-button  @click="showModel" plain>主要按钮</el-button>
            </el-col> -->
        </el-row>
        <div>
          <el-row>
            <el-col :span="12" v-show="show">
              <el-table :data="tableData" style="width: 100%" max-height="height" border>
                <!-- <el-table-column prop="订单号" label="订单号" width="180">
                </el-table-column>
                <el-table-column prop="客户编号" label="客户编号">
                </el-table-column>
                <el-table-column prop="客户姓名" label="客户姓名">
                </el-table-column> -->
                <el-table-column v-for="(value, key, index) in tableData[0]" :key="index" :prop="key"
                  :label="key"></el-table-column>
              </el-table>
            </el-col>
            <el-col :span="12" v-show="!show"><el-skeleton /></el-col>

            <el-col :span="12" max-height="height">
              <div id="wordView" v-html="wordText"></div>
            </el-col>
            <!-- <el-col :span="1"> </el-col>
            <el-col :span="6">
              <el-skeleton />
            </el-col> -->
          </el-row>
        </div>
        <div>
          <!-- <el-select v-model="value" placeholder="请选择">
            <el-option
              v-for="item in options"
              :key="item.value"
              :label="item.label"
              :value="item.value"
            >
            </el-option>
          </el-select> -->
        </div>

      </el-main>
    </el-container>
  </div>
</template>

<script>
import { read, utils, writeFile } from "xlsx";
import { readFile, character, delay } from "../assets/lib/utils";
import { Loading } from "element-ui";
import docxtemplater from "docxtemplater";
//import JSZip from "jszip";
import PizZip from "pizzip";
import JSZipUtils from "jszip-utils";
import { saveAs } from "file-saver";
import mammoth from "mammoth";

export default {
  name: "Hello",
  data() {
    return {
      height: document.documentElement.clientHeight - 130,
      show: false,
      fileName: "",
      fileList: [],
      tableData: [],
      wordText: "",
      wordURL: "",
      filesPath:"",
      modelUrl:"model_1.docx",
      options: [
        {
          value: "选项1",
          label: "黄金糕",
        }
      ],
      value: "",
    };
  },
  methods: {
    submitUpload() {
      this.$refs.upload.submit();
    },
    handleRemove(file, fileList) {
      console.log(file, fileList);
      this.show = false;
      this.tableData = [];
    },
    handlePreview(file) {
      console.log(file);
    },
    async handle(ev) {
      //console.log(ev);
      let file = ev.raw;
      if (!file) return;
      console.log(file);
      this.fileName = file.name;

      this.show = false;
      let loadingInstance = Loading.service({
        text: "loading...",
      });
      await delay(100);

      //读取file数据转json
      let data = await readFile(file);
      let workbook = read(data, { type: "binary" }),
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
      data = utils.sheet_to_json(worksheet);

      //处理数据
      let arr = [];
      data.forEach((item) => {
        let obj = {};
        if (character.length > 0) {
          for (let key in character) {
            if (!character.hasOwnProperty(key)) break;
            let v = character[key],
              text = v.text,
              type = v.type;
            type === "string" ? (v = String(v)) : null;
            type === "number" ? (v = Number(v)) : null;
            v = item[text] || "";

            obj[key] = v;
          }
        } else {
          obj = item;
        }

        arr.push(obj);
      });

      await delay(100);
      this.show = true;
      this.tableData = arr;
      console.log(this.tableData);
      loadingInstance.close();
    },
    exportWord() {
      let that = this;
      if (this.tableData.length === 0) return;

      // 读取并获得模板文件的二进制内容
      JSZipUtils.getBinaryContent(that.modelUrl, function (error, content) {
        // model.docx是模板。我们在导出的时候，会根据此模板来导出对应的数据
        // 抛出异常
        if (error) {
          throw error;
        }

        //let zip = new JSZip();
        //zip.file(content);

        // 创建一个PizZip实例，内容为模板的内容
        let zip = new PizZip(content);
        // 创建并加载docxtemplater实例对象
        let doc = new docxtemplater().loadZip(zip);
        doc.setOptions({
          nullGetter: function () {
            return "";
          },
        });
        // 设置模板变量的值
        let tempData = {};
        tempData.datas = that.tableData;
        //let temp  = JSON.stringify(tempData);
        console.log(tempData);
        doc.setData({
          datas: that.tableData,
        });
        console.log(that.tableData);
        console.log(doc);
        try {
          // 用模板变量的值替换所有模板变量
          doc.render();
        } catch (error) {
          // 抛出异常
          let e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
          };
          console.log(JSON.stringify({ error: e }));
          throw error;
        }

        // 生成一个代表docxtemplater对象的zip文件（不是一个真实的文件，而是在内存中的表示）
        let out = doc.getZip().generate({
          type: "blob",
          mimeType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        // 将目标文件对象保存为目标类型的文件，并命名
        saveAs(
          out,
          that.fileName.replace(".xlsx", "").replace(".xls", "") +
          "[done]" +
          ".docx"
        );
      });
    },
    showWord() {
      mammoth.convertToHtml({ path: "path/to/document.docx" })
        .then(function (result) {
          var html = result.value; // The generated HTML
          var messages = result.messages; // Any messages, such as warnings during conversion
        })
        .catch(function (error) {
          console.error(error);
        });
    },
    //选择本地文件预览
    changeFile(event) {
      let that = this;
      console.log(event);
      let file = event.target.files[0];
      console.log(file);
      if (!file) return;

      let reader = new FileReader();
      reader.onload = function (loadEvent) {
        console.log(loadEvent);
        let arrayBuffer = new Uint8Array(loadEvent.target.result); //arrayBuffer
        //let arrayBuffer =  loadEvent.target.result; //arrayBuffer
        mammoth
          .convertToHtml({ arrayBuffer: arrayBuffer })
          //.then(that.displayResult)
          .then(function (resultObject) {
            that.$nextTick(() => {
              //console.log(resultObject.value);
              that.wordText = resultObject.value
            })
          })
          .catch(function (error) {
            console.error(error);
          });
      };
      reader.readAsArrayBuffer(file);

    },
    //页面渲染
    displayResult(result) {
      this.wordText = result.value;
    },
    //调用后台文档下载接口
    getWordText() {
      console.log(this.wordURL);
      let that = this;
      const xhr = new XMLHttpRequest();
      xhr.open("get", "http://127.0.0.1:3000/downFile", true);
      xhr.responseType = "arraybuffer";
      xhr.onload = function () {
        if (xhr.status == 200) {
          mammoth
            .convertToHtml({ arrayBuffer: new Uint8Array(xhr.response) })
            .then(that.displayResult);
        }
      };
      xhr.send();
    },
    changeFiles(event) {
      console.log(event);
    },
    showModel(){
      if(this.filesPath==="")return;
      let path  = require(this.filesPath);
      console.log(path);
    }
  },
};
</script>

<style>

</style>
