import { readFile, utils } from 'xlsx';

// This will be called when the files change
const parseSheet = (file, { path }) => {
  // `path` is a path that can be read with `XLSX.readFile`
  const wb = readFile(path);
  const o = wb.SheetNames.map(name => ({ name, data: utils.sheet_to_json(wb.Sheets[name])}));
  return { data: o };
}

export default {
  // Disable server-side rendering: https://go.nuxtjs.dev/ssr-mode
  ssr: false,

  // Global page headers: https://go.nuxtjs.dev/config-head
  head: {
    title: 'anuo_vue',
    htmlAttrs: {
      lang: 'en'
    },
    meta: [
      { charset: 'utf-8' },
      { name: 'viewport', content: 'width=device-width, initial-scale=1' },
      { hid: 'description', name: 'description', content: '' },
      { name: 'format-detection', content: 'telephone=no' }
    ],
    link: [
      { rel: 'icon', type: 'image/x-icon', href: '/favicon.ico' }
    ]
  },

  // Global CSS: https://go.nuxtjs.dev/config-css
  css: [
  ],

  // Plugins to run before rendering page: https://go.nuxtjs.dev/config-plugins
  plugins: [
  ],

  // Auto import components: https://go.nuxtjs.dev/config-components
  components: true,

  // Modules for dev and build (recommended): https://go.nuxtjs.dev/config-modules
  buildModules: [
  ],

  // Modules: https://go.nuxtjs.dev/config-modules
  modules: [
    // https://go.nuxtjs.dev/bootstrap
    'bootstrap-vue/nuxt',
  ],

  // Build Configuration: https://go.nuxtjs.dev/config-build
  build: {
  },
  // content.extendParser allows us to hook into the parsing step
  content: {
    extendParser: {
      // the keys are the extensions that will be matched.  The "." is required
      ".numbers": parseSheet,
      ".xlsx": parseSheet,
      ".xls": parseSheet,
      // can add other extensions like ".fods" as desired
    }
  },
}
