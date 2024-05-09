<template>
    <el-card class="box-card">
      <div slot="header" class="clearfix">
        <span>N1催收结算</span>
      </div>
  
      <el-button type="primary" @click="upload">上传表格</el-button>
      <el-button type="success" @click="handleData">处理数据</el-button>
      <el-button type="info" @click="download">下载表格</el-button>
    </el-card>
  </template>
  
  <script>
  export default {
    name: 'N1Settlement',
    methods: {
      upload() {
        // 上传表格的逻辑
      },
      handleData() {
        // 处理数据的逻辑
      },
      download() {
        // 下载表格的逻辑
      }
    }
  }
  </script>
  
  <style scoped>
  .box-card {
    width: 100%;
    margin: 30px auto;
  }
  
  .el-button {
    margin: 0 10px 10px 0;
  }
  </style>