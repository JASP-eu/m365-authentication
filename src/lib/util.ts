export const getSearchParams = (hash: string) => {
  const params: any = {}
  const vars = hash.slice(1).split('&')

  for (var i = 0; i < vars.length; i++) {
    var pair = vars[i].split('=')
    params[pair[0]] = decodeURIComponent(pair[1])
  }

  return params
}

const _window: any = window
export const isApp = () => !!_window.forge
export const openWithOptions = _window.forge?.tabs?.openWithOptions ?? (() => undefined)

const ua = window.navigator.userAgent
const msie = ua.indexOf('MSIE ')
const msie11 = ua.indexOf('Trident/')
const msEdge = ua.indexOf('Edge/')
export const isIE = msie > 0 || msie11 > 0
export const isEdge = msEdge > 0
