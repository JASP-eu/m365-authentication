import pkceChallenge from 'pkce-challenge'

import { getDefaultSharePointTenant } from './sharePoint'
import { getSearchParams, isApp, isIE, isEdge, openWithOptions } from './util'
import {
  POPUP,
  RESOLVING_REDIRECT_TOKEN,
  MS365_STATUS,
  MS365_CODE,
  MS365_REFRESH_TOKEN,
  MS365_VERIFIER,
  RESOLVING_POPUP_TOKEN,
  REDIRECT,
  MS365_ACCESS_TOKEN,
  MS365_ID_TOKEN,
  MS365_SIGN_IN_CANCELLED_EVENT,
  MS365_SIGN_IN_EVENT,
  MS365_SIGN_OUT_EVENT,
  MS365_TOKEN_EXPIRY_TIMESTAMP,
} from './constants'

const mobileOriginUrl = 'https://localhost:9360/src/index.html'

export class AuthProvider {
  /**
   * This code runs on startup. It checks if we're currently on a redirect page.
   * If so, it extracts the code it's been passed and saves it to the LocalStorage.
   *
   * If called by `loginPopup`, saving the code triggers the handler code of the calling window.
   *
   * Otherwise, if called by `loginRedirect`, the code is used right here to acquire tokens.
   * Then, the `?code=...` appendix is removed.
   */
  constructor(clientId: string, scopes: string[], redirectUri: string) {
    this.clientId = clientId
    this.scopes = scopes
    this.redirectUri = redirectUri

    if (typeof window !== 'undefined') {
      const codeIndex = window.location.href.indexOf('?code=')
      if (codeIndex) {
        const code = getSearchParams(window.location.href.slice(codeIndex)).code
        if (code) {
          if (this.status === POPUP) {
            this.code = code
            window.close()
          } else {
            this.status = RESOLVING_REDIRECT_TOKEN

            try {
              this.acquireTokenViaCode(code)
            } catch (e) {
              // invalid state produced by user
            }

            const href = window.location.href
            const urlWithoutHash = href.slice(0, href.indexOf('?code='))
            window.history.pushState({}, document.title, urlWithoutHash)
          }
        } else {
          this.clearTempAuthState()
        }
      }
    }
  }

  private clientId: string
  private scopes: string[]
  private redirectUri: string

  private challenge = pkceChallenge()

  private get status() {
    return localStorage.getItem(MS365_STATUS) ?? ''
  }

  private set status(token: string) {
    localStorage.setItem(MS365_STATUS, token)
  }

  private get code() {
    return localStorage.getItem(MS365_CODE) ?? ''
  }

  private set code(code: string) {
    localStorage.setItem(MS365_CODE, code)
  }

  get idToken() {
    const token = localStorage.getItem(MS365_ID_TOKEN) ?? ''

    if ((!token || token === 'undefined') && !this.status) {
      this.logout()
    }

    return token
  }

  set idToken(token: string) {
    localStorage.setItem(MS365_ID_TOKEN, token)
  }

  get accessToken() {
    return localStorage.getItem(MS365_ACCESS_TOKEN) ?? ''
  }

  set accessToken(token: string) {
    localStorage.setItem(MS365_ACCESS_TOKEN, token)
  }

  private get refreshToken() {
    return localStorage.getItem(MS365_REFRESH_TOKEN) ?? ''
  }

  private set refreshToken(token: string) {
    localStorage.setItem(MS365_REFRESH_TOKEN, token)
  }

  private get expiryTimestamp() {
    return parseInt(localStorage.getItem(MS365_TOKEN_EXPIRY_TIMESTAMP) ?? '0', 10)
  }

  private set expiryTimestamp(timestampInSeconds: number) {
    localStorage.setItem(MS365_TOKEN_EXPIRY_TIMESTAMP, timestampInSeconds.toString())
  }

  private get verifier() {
    return localStorage.getItem(MS365_VERIFIER) ?? ''
  }

  private set verifier(verifier: string) {
    localStorage.setItem(MS365_VERIFIER, verifier)
  }

  /**
   * Generates the url to Microsoft's /authorize endpoint.
   *
   * @returns the url
   */
  getAuthCodeLoginUrl() {
    const challenge = pkceChallenge()

    return (
      'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?' +
      `client_id=${this.clientId}` +
      `&scope=${encodeURIComponent(this.scopes.join(' '))}` +
      `&redirect_uri=${encodeURIComponent(this.redirectUri)}` +
      `&state=${encodeURIComponent(`originUrl=[${window.location.href}]`)}` +
      '&response_type=code' +
      '&response_mode=query' +
      `&code_challenge=${challenge.code_challenge}` +
      '&code_challenge_method=S256'
    )
  }
  /**
   * Looks up the default SharePoint Tenant Url from MS Graph.
   * Then acquires a token specifically for reading sites on that SharePoint
   */
  connectToSharePoint = async (sharepointServerUrl?: string) => {
    if (!sharepointServerUrl) {
      // get sharepoint server url
      const sharepointTenant = await getDefaultSharePointTenant()
      sharepointServerUrl = sharepointTenant.url
    }

    const xhr = new XMLHttpRequest()
    xhr.open('POST', 'https://login.microsoftonline.com/common/oauth2/v2.0/token', true)
    xhr.setRequestHeader('Access-Control-Allow-Origin', window.location.origin)
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded')
    const tokenResult = new Promise((resolve: (value?: string) => any) => {
      xhr.onreadystatechange = () => {
        if (xhr.readyState === 4) {
          if (xhr.response) {
            const response = JSON.parse(xhr.response)

            if (response.error) {
              resolve(undefined)
            } else {
              resolve(response.access_token)
            }
          }
        }
      }
    })
    xhr.send(
      `client_id=${this.clientId}` +
        `&scope=${encodeURIComponent(sharepointServerUrl + '/Sites.Read.All')}` +
        `&refresh_token=${this.refreshToken}` +
        `&grant_type=refresh_token`
    )

    const token = await tokenResult
    if (!token) {
      throw new Error('Could not connect to SharePoint')
    }

    return {
      url: sharepointServerUrl,
      token,
    }
  }

  /**
   * Loads the current access token.
   * If token is expired or about to expire in less than 5 minutes, refreshes it first.
   */
  async getValidAccessToken() {
    if (Math.floor(new Date().valueOf() / 1000) - 300 >= this.expiryTimestamp) {
      await this.refreshAccessToken()
    }

    return this.accessToken
  }

  /**
   * ENTRY POINT
   * When coming from the WithAuthentication wrapper, this is where we start the auth process.
   *
   * Depending on the device and browser, calls the appropriate login function.
   * Uses Forge for mobile devices, Redirect flow for IE11 and legacy Edge, Popup flow otherwise.
   */
  login() {
    if (isApp()) {
      this.loginViaForge()
    } else if (isIE || isEdge) {
      this.loginRedirect()
    } else {
      const popup = this.loginPopup()

      if (!popup) {
        console.warn("Couldn't open popup, trying redirect instead.")
        this.loginRedirect()
      }

      let interval: number

      const checkPopupClosed = () => {
        if (popup?.closed) {
          window.dispatchEvent(
            new Event(MS365_SIGN_IN_CANCELLED_EVENT, { bubbles: true, cancelable: false })
          )

          clearInterval(interval)
        }
      }

      interval = window.setInterval(checkPopupClosed, 300)
    }
  }

  /**
   * Request new refresh and access tokens from login.microsoftonline.com
   * Then persists them in the localStorage.
   */
  async refreshAccessToken(newScopes: string[] = this.scopes) {
    if (this.refreshToken) {
      try {
        const xhr = new XMLHttpRequest()
        xhr.open('POST', 'https://login.microsoftonline.com/common/oauth2/v2.0/token', true)
        xhr.setRequestHeader('Access-Control-Allow-Origin', window.location.origin)
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded')
        xhr.onreadystatechange = () => {
          if (xhr.readyState === 4) {
            if (xhr.response) {
              const response = JSON.parse(xhr.response)
              const currentTimestamp = Math.floor(new Date().valueOf() / 1000)

              if (response.error) {
                console.error('access token refresh:', response.error)
                this.logout()
              } else {
                this.idToken = response.id_token
                this.refreshToken = response.refresh_token
                this.accessToken = response.access_token
                this.expiryTimestamp = currentTimestamp + response.expires_in
              }
            }
          }
        }
        xhr.send(
          `client_id=${this.clientId}` +
            `&scope=${encodeURIComponent(newScopes.join(' '))}` +
            `&refresh_token=${this.refreshToken}` +
            `&grant_type=refresh_token`
        )
      } catch (e) {
        console.warn('Refresh Token expired. Trying new Login.')
        this.login()
      }
    } else {
      this.logout()
    }
  }

  /**
   * Used to register an action to execute upon successful sign-in.
   */
  doOnSignIn(func: () => any) {
    window.addEventListener(MS365_SIGN_IN_EVENT, func, { once: true })
    return () => window.removeEventListener(MS365_SIGN_IN_EVENT, func)
  }

  /**
   * Used to register an action to execute upon a cancelled sign-in process.
   */
  doOnSignInCancelled(func: () => any) {
    window.addEventListener(MS365_SIGN_IN_CANCELLED_EVENT, func, { once: true })
    return () => window.removeEventListener(MS365_SIGN_IN_CANCELLED_EVENT, func)
  }

  /**
   * Used to register an action to execute upon sign out.
   */
  doOnSignOut(func: () => any) {
    window.addEventListener(MS365_SIGN_OUT_EVENT, func, { once: true })
    return () => window.removeEventListener(MS365_SIGN_OUT_EVENT, func)
  }

  /**
   * Removes all traces of an ongoing authentication process as well as an authenticated session.
   * Then communicates this logout to all other interested parties on the page.
   */
  logout(broadcast = true) {
    this.clearTempAuthState()
    localStorage.removeItem(MS365_ID_TOKEN)
    localStorage.removeItem(MS365_REFRESH_TOKEN)
    localStorage.removeItem(MS365_ACCESS_TOKEN)
    localStorage.removeItem(MS365_TOKEN_EXPIRY_TIMESTAMP)

    if (broadcast) {
      window.dispatchEvent(new Event(MS365_SIGN_OUT_EVENT, { bubbles: true, cancelable: false }))
    }
  }

  /**
   * Removes all traces of an ongoing authentication.
   * Used if the process has been cancelled or completed.
   */
  private clearTempAuthState() {
    localStorage.removeItem(MS365_STATUS)
    localStorage.removeItem(MS365_CODE)
    localStorage.removeItem(MS365_VERIFIER)
  }

  /**
   * This is the first request after the initial handshake where we got a code.
   * Using this code, it requests an access token as well as a refresh token.
   * Usually called by the constructor right after the page is opened.
   */
  private acquireTokenViaCode(code: string) {
    const verifier =
      this.status === RESOLVING_REDIRECT_TOKEN ? this.verifier : this.challenge.code_verifier

    const xhr = new XMLHttpRequest()
    xhr.open('POST', 'https://login.microsoftonline.com/common/oauth2/v2.0/token', true)
    xhr.setRequestHeader('Access-Control-Allow-Origin', window.location.origin)
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded')
    xhr.onreadystatechange = () => {
      if (xhr.readyState === 4) {
        if (xhr.response) {
          if (typeof Storage !== 'undefined') {
            const response = JSON.parse(xhr.response)

            if (!response.error) {
              this.clearTempAuthState()
              this.idToken = response.id_token
              this.refreshToken = response.refresh_token
              this.accessToken = response.access_token
              this.expiryTimestamp = Math.floor(new Date().valueOf() / 1000) + response.expires_in

              window.dispatchEvent(
                new Event(MS365_SIGN_IN_EVENT, { bubbles: true, cancelable: false })
              )
            } else {
              // invalid grant. we weren't expecting to receive the token
              this.clearTempAuthState()
            }
          } else {
            // no storage
            this.clearTempAuthState()
          }
        } else {
          // no response
          this.clearTempAuthState()
        }
      } else {
        // request failed
        this.clearTempAuthState()
      }
    }

    xhr.send(
      `client_id=${this.clientId}` +
        `&scope=${encodeURIComponent(this.scopes.join(' '))}` +
        `&redirect_uri=${encodeURIComponent(this.redirectUri)}` +
        // `&state=${encodeURIComponent(`originUrl=[${window.location.href}]`)}` +
        `&code=${code}` +
        `&grant_type=authorization_code` +
        `&code_verifier=${verifier}`
    )
  }

  /**
   * Opens a popup calling login.microsoftonline.com, which then redirects to /AuthRedirect/Index
   * The constructor of this class will be called on that AuthRedirect page.
   * There it'll be closed, as well.
   */
  private loginPopup() {
    this.status = POPUP

    const storageListener = () => {
      if (this.status === POPUP) {
        if (this.code) {
          this.status = RESOLVING_POPUP_TOKEN
          this.acquireTokenViaCode(this.code)

          window.removeEventListener('storage', storageListener)
        }
      }
    }
    window.addEventListener('storage', storageListener)

    return window.open(
      this.getAuthCodeLoginUrl(),
      undefined,
      'status=no,location=no,toolbar=no,menubar=no,width=400,height=600'
    )
  }

  /**
   * Redirects to login.microsoftonline.com and then back here.
   * The constructor of this class will be called again and acquire the access token via the code.
   */
  private loginRedirect() {
    this.status = REDIRECT
    this.verifier = this.challenge.code_verifier

    window.location.replace(this.getAuthCodeLoginUrl())
  }

  /**
   * Uses Forge's `openWithOptions` function to open a popup mimicking the `loginPopup` logic.
   */
  private loginViaForge() {
    this.status = POPUP

    openWithOptions(
      {
        url: this.getAuthCodeLoginUrl(),
        pattern: `${mobileOriginUrl}*`,
      },
      (data: { url: string }) => {
        // TODO: won't this be done by the constructor, anyway?
        const search = data.url.slice(data.url.indexOf('?code='))
        const code = getSearchParams(search).code
        if (code) {
          try {
            this.status = RESOLVING_POPUP_TOKEN
            this.acquireTokenViaCode(code)
          } catch (e) {
            this.clearTempAuthState()
            console.error(e.message)
          }
        } else {
          // user cancelled
          this.clearTempAuthState()
        }
      }
    )
  }
}
