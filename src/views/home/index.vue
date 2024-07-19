<template>
  <div class="app-container">
    <div>
      <el-button class="button" icon="el-icon-s-promotion" @click="goToRoute('/stockOrder')">
        前往备货单
      </el-button>
    </div>
    <div>
      <el-button class="button" icon="el-icon-s-promotion" @click="goToRoute('/mergeTable')">
        前往合并表格
      </el-button>
    </div>
    <div>
      <el-button class="button" icon="el-icon-s-promotion" @click="goToRoute('/orderMerge')">
        前往订单合并
      </el-button>
    </div>
  </div>
</template>

<script>
export default {
  data() {
    return {}
  },
  methods: {
    goToRoute(path) {
      this.$router.push({
        path
      })
    }
  }
}
</script>

<style>
.app-container {
  display: flex;
  flex-direction: column;
  gap: 10px;

  .button {
    flex: 1;
    width: 200px;
  }
}
</style>