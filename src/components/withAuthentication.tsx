import React, { useState, useEffect, useCallback } from 'react'
import jwtDecode from 'jwt-decode'
import { addAuthorizationToken } from '@jasp/m365-graph'

import { ErrorBoundary } from '../components/ErrorBoundary'
import { AuthProvider } from '../lib/authProvider'

interface AccountInfo {
  environment: string
  username: string
  homeAccountId: string
  tenantId: string
}

export interface AuthenticationProps {
  account: AccountInfo
  login: () => void
  logout: () => void
}

export interface LoginProps {
  login: () => void
}

export const withAuthentication = function <T>(
  WrappedComponent: React.ComponentType<T & AuthenticationProps>,
  LoginComponent: React.ComponentType<LoginProps>,
  LoadingComponent: React.ComponentType,
  authProvider: AuthProvider
) {
  return (props: T) => {
    const [account, setAccount] = useState<AccountInfo>()
    const [isLoading, toggleLoading] = useState(true)
    const [error, setError] = useState(false)

    const login = () => {
      toggleLoading(true)
      setError(false)
      authProvider.login()
    }

    const logout = useCallback((broadcast = true) => {
      setError(false)
      setAccount(undefined)
      authProvider.logout(broadcast)
    }, [])

    /**
     * use the idToken from the authProvider to retrieve account data
     */
    const getAccount = useCallback(() => {
      try {
        if (authProvider.idToken) {
          const decryptedToken: {
            name: string // John Doe
            preferred_username: string // john.doe@example.com
            oid: string // accountId
            tid: string // tenantId
          } = jwtDecode(authProvider.idToken)

          const newAccountInfo: AccountInfo = {
            environment: 'login.microsoftonline.com',
            username: decryptedToken.preferred_username,
            homeAccountId: `${decryptedToken.oid}.${decryptedToken.tid}`,
            tenantId: decryptedToken.tid,
          }

          setAccount(newAccountInfo)
        } else {
          toggleLoading(false)
        }
      } catch (e) {
        console.error(`getAccount: `, e)
        setError(true)
      }
    }, [])

    /**
     * retrieve access token and configure future graph requests
     */
    const ensureToken = useCallback(async () => {
      try {
        const accessToken = await authProvider.getValidAccessToken()
        addAuthorizationToken(accessToken)

        toggleLoading(false)
      } catch (e) {
        console.error(`ensureToken: `, e)
      }
    }, [])

    /**
     * load the signed-in user's account data
     */
    useEffect(() => {
      getAccount()
    }, [getAccount])

    /**
     * as soon as we're signed in, configure the graph requests
     */
    useEffect(() => {
      if (account) {
        ensureToken()
      }
    }, [account, ensureToken])

    /**
     * handle if user has signed in elsewhere
     */
    useEffect(() => {
      if (!account) {
        return authProvider.doOnSignIn(getAccount)
      }
    }, [account, getAccount])

    /**
     * handle if user cancelled the sign-in popup
     */
    useEffect(() => {
      if (!account) {
        return authProvider.doOnSignInCancelled(() => toggleLoading(false))
      }
    }, [account, getAccount])

    /**
     * handle if user has signed out elsewhere
     */
    useEffect(() => {
      if (account) {
        return authProvider.doOnSignOut(() => logout(false))
      }
    }, [account, logout])

    /**
     * render function
     */
    const getContent = () => {
      if (isLoading) {
        return <LoadingComponent />
      }

      if (account) {
        return <WrappedComponent account={account} login={login} logout={logout} {...props} />
      }

      if (error) {
        return <pre>{error}</pre>
      }

      return <LoginComponent login={login} />
    }

    return <ErrorBoundary>{getContent()}</ErrorBoundary>
  }
}
