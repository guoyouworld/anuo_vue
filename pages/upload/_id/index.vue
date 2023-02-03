<template>
  <div>
    <h1 class="red">Hello {{ this.$route.params.id }}!</h1>
  </div>
</template>
<script>
export default {
  validate({ params }) {
    // 必须是number类型
    return /^\d+$/.test(params.id)
  }
}
</script>
