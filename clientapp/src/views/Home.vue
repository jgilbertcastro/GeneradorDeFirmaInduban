<template>
    <div class="home">
        <CreateSigning msg="Welcome to Your Vue.js App"/> 
    </div>
</template>

<script>
 // @ is an alias to /src
    import CreateSigning from '@/components/CreateSigning.vue'

export default {
  name: 'Home',
  components: {
      CreateSigning
  }
};
</script>
