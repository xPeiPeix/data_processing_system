<script setup lang="ts">
import { IconVuejs, IconJs } from '@iconify-prerendered/vue-fa-brands';
import { IconCardsHeart,IconAlphaXCircle } from '@iconify-prerendered/vue-mdi';
</script>
<template>
  <p>
    <IconJs style="color: #efd81d" />
    <IconCardsHeart style="color: #f03a17" />
    <IconVuejs style="color: #3fb27f" />
    <IconAlphaXCircle style="color: #3fb27f"/>
  </p>
</template>

<style scoped>
p {
  display: flex;
  gap: 1rem;
  font-size: 100px;
}
</style>
