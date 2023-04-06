<template>
  <q-page>
    <q-stepper
      v-model="step"
      ref="stepper"
    >
      <q-step
        :name="1"
        :done="step > 1"
        title="导入文件"
        icon="text_snippet"
      >
        <q-form
          @submit="onSubmit"
          @reset="onReset"
        >
          <q-input
            label="正则文件"
            readonly
            lazy-rules
            :rules="[ val => val && val.length > 0 || '请选择正则文件']"
            v-model="filePathForm.regFilePath"
          >
            <template v-slot:before>
              <q-icon name="file_copy" />
            </template>
            <template v-slot:append>
              <q-btn round dense flat icon="content_paste_search" color="primary" @click="chooseFilePath('regFilePath')"/>
            </template>
          </q-input>
          <q-input
            label="内容文件"
            readonly
            lazy-rules
            v-model="filePathForm.contentFilePath"
            :rules="[ val => val && val.length > 0 || '请选择内容文件']"
          >
            <template v-slot:before>
              <q-icon name="file_copy" />
            </template>
            <template v-slot:append>
              <q-btn round dense flat icon="content_paste_search" color="primary" @click="chooseFilePath('contentFilePath')"/>
            </template>
          </q-input>
          <div class="q-pt-md">
            <q-btn label="确认" type="submit" color="primary"/>
            <q-btn label="重置" type="reset" color="primary" flat class="q-ml-sm" />
          </div>
        </q-form>
      </q-step>
      <q-step
        :name="2"
        :done="step > 2"
        title="填写规则"
        icon="edit_note"
      >
        <q-form
          @submit="executeMatch"
        >
          <q-select
            v-model="contentColumnSelect"
            :options="contentColumns"
            :rules="[ val => !!val || '请选择文本列' ]"
            label="文本列"
            lazy-rules
          />
          <q-select
            v-model="regColumnSelect"
            :options="regColumns"
            label="正则描述列"
            :rules="[ val => !!val || '请选择正则描述列' ]"
            lazy-rules
          />
          <q-field label="正则列展示" stack-label borderless >
            <template v-slot:control>
              <ol start="0">
                <li class="q-mb-sm" v-for="(item, index) in regColumns" :key="index">{{item}}</li>
              </ol>
            </template>
          </q-field>
          <q-toggle
            v-model="isNotCaseSensitive"
            label="不区分大小写"
            color="green"
          />
          <q-toggle
            v-model="isFilterEmoticon"
            label="剔除[表情]"
            color="orange"
          />
          <q-toggle
            v-model="isFilterURL"
            label="剔除链接"
            color="pink"
          />
          <q-toggle
            v-model="isFilterUserName"
            label="剔除@用户名"
            color="purple" l
          />
          <q-toggle
            v-model="isFilterTopic"
            label="剔除#话题#"
            color="lightBlue"
            @update:model-value="val => topicFilterChange(val, 'isFilterTopic')"
          />
          <q-toggle
            v-model="isFilterSpecTopic"
            label="剔除#特殊格式话题 "
            color="red"
            @update:model-value="val => topicFilterChange(val, 'isFilterSpecTopic')"
          />
          <q-toggle
            v-model="isSplit"
            label="启用分割字符"
            color="primary"
          />
          <q-input
            v-if="isSplit"
            v-model.trim="splitText"
            label="分割字符(正则表达式)"
            hint="若为空，默认使用，|。"
          />
          <q-input
            v-model.trim="logicCode"
            label="逻辑码表"
            lazy-rules
            :rules="[
              val => !!val || '逻辑码表不能为空',
              val => validLogicCode(val) || '语法错误，请检查。列如: 2 && !(3 && 4)',
              val => validLogicCodeIndex(val) || '数字索引超出范围'
            ]"
            hint="请使用数字索引、英文括号、“&&”表示“与”，“||”表示“或”，“!”表示“非”"
          />
          <div class="q-pt-md">
            <q-btn label="确认" type="submit" color="primary"/>
            <q-btn label="上一步" type="reset" color="primary" flat class="q-ml-sm" @click="backStep"/>
          </div>
        </q-form>
      </q-step>
      <q-step
        :name="3"
        title="执行匹配"
        icon="smart_toy"
      >
        <div class="row items-center justify-evenly">
          <div class="result-wrapper">
            <q-circular-progress
              show-value
              font-size="24px"
              :value="progressValue"
              size="120px"
              :thickness="0.22"
              color="teal"
              track-color="grey-3"
              class="q-ma-md"
            >
              {{ progressValue }}%
            </q-circular-progress>
            <q-item-label v-if="progressFinished"> 执行完毕，请到内容所在目录查看执行结果！ </q-item-label>
            <q-item-label v-else> 执行匹配中请稍后... </q-item-label>
            <q-btn v-if="progressFinished" label="上一步" type="reset" color="primary" flat class="q-ml-sm" @click="backStep"/>
          </div>
        </div>
      </q-step>
    </q-stepper>
  </q-page>
</template>

<script setup lang="ts">
import { ref, reactive } from 'vue'
import { QStepper } from 'quasar';

const step = ref(1);
const isSplit = ref(true);
const isFilterUserName = ref(true);
const isFilterTopic = ref(true);
const isFilterSpecTopic = ref(false);
const isNotCaseSensitive = ref(true);
const isFilterEmoticon = ref(true);
const isFilterURL = ref(true);
const splitText = ref('，|。');
const logicCode = ref('');

// 话题过滤器变换
const topicFilterChange = (val:boolean, filed:string) => {
  if (val) {
    if(filed === 'isFilterTopic') {
      isFilterSpecTopic.value = false;
    } else {
      isFilterTopic.value = false;
    }
  }
}

//校验语法正确
const validLogicCode = (val: string) => {
  const replaceCodeStr = val.replace(/\d+/g, 'true');
  try {
    const pass = eval(replaceCodeStr);
    return typeof pass === 'boolean';
  } catch (e) {
    return false;
  }
}

//校验是否索引越界
const validLogicCodeIndex = (val: string) => {
  const regNumList = val.match(/\d+/g); //对逻辑码表的数字进行切分
  if (regNumList) {
    return regNumList.every(item => parseInt(item) < regColumns.value.length);
  } else {
    return true;
  }
}

const filePathForm = reactive({
  regFilePath: '',
  contentFilePath: ''
});

const stepper = ref<QStepper | null>(null);

const chooseFilePath = async (filePathField: keyof typeof filePathForm) => {
  const filePath = await window.electronAPI.openFile();
  if (filePath) {
    filePathForm[filePathField] = filePath;
  }
}

const contentColumns = ref([]);
const regColumns = ref([]);
const contentColumnSelect = ref();
const regColumnSelect = ref();

const onSubmit = async () => {
  stepper.value?.next();
  regColumnSelect.value = null;
  contentColumnSelect.value = null;
  contentColumns.value = await window.electronAPI.parseColumns(filePathForm.contentFilePath);
  regColumns.value = await window.electronAPI.parseColumns(filePathForm.regFilePath);
}


const onReset = () => {
  filePathForm.regFilePath = '';
  filePathForm.contentFilePath = '';
}

/**
 * @description 上一步
 */
const backStep = () => {
  stepper.value?.previous();
}

const progressValue = ref(0);
const progressFinished = ref(false);

/**
 * @description 执行批量正则匹配
 */
const executeMatch = () => {
  progressValue.value = 0;
  progressFinished.value = false;
  stepper.value?.next();
  window.electronAPI.executeRegMatch({
    regFilePath: filePathForm.regFilePath,
    contentFilePath: filePathForm.contentFilePath,
    regColumnSelect: regColumns.value.findIndex(value => regColumnSelect.value === value),
    contentColumnSelect: contentColumns.value.findIndex(value => contentColumnSelect.value === value),
    logicCode: logicCode.value,
    splitText: isSplit.value ? splitText.value || '，|。' : '',
    isFilterUserName: isFilterUserName.value,
    isFilterTopic: isFilterTopic.value,
    isFilterSpecTopic: isFilterSpecTopic.value,
    isNotCaseSensitive: isNotCaseSensitive.value,
    isFilterEmoticon: isFilterEmoticon.value,
    isFilterURL: isFilterURL.value,
  })
}

window.electronAPI.onUpdateState((event, res) => {
  if (res.state === 'done') {
    progressFinished.value = true;
  } else {
    progressValue.value = Number.parseFloat(res.progress);
  }
})
</script>

<style scoped>
.result-wrapper {
  display: flex;
  flex-flow: column;
  justify-content: center;
  align-items: center;
}
</style>
