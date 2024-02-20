<template>
  <q-layout view="lHh Lpr lFf">
    <q-header elevated>
      <q-toolbar>
        <q-btn
          flat
          dense
          round
          icon="menu"
          aria-label="Menu"
          @click="toggleLeftDrawer"
        />

        <q-toolbar-title> 文本智能识别工具 </q-toolbar-title>
      </q-toolbar>
    </q-header>

<!--    <q-drawer v-model="leftDrawerOpen" bordered>-->
<!--      <q-list>-->
<!--        <q-item-label header> 支持作者 </q-item-label>-->
<!--        <q-img-->
<!--          class="q-ml-xl shadow-8"-->
<!--          src="~assets/wechat_pay_qr.png"-->
<!--          :ratio="1"-->
<!--          width="70%"-->
<!--        />-->
<!--        <q-item-label class="q-mt-lg pay-author-label"> 请作者喝杯奶茶 </q-item-label>-->
<!--      </q-list>-->
<!--    </q-drawer>-->

    <q-page-container>
      <router-view />
    </q-page-container>
  </q-layout>
</template>

<script setup lang="ts">
import { ref } from 'vue';

const leftDrawerOpen = ref(false);

function toggleLeftDrawer() {
  leftDrawerOpen.value = !leftDrawerOpen.value;
}
</script>

<style scoped>
.pay-author-label {
  text-align: center;
  color: #1976d2;
}
</style>
