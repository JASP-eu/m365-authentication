import { Site } from '@microsoft/microsoft-graph-types'
import { getGraph } from '@jasp/m365-graph'

export const getDefaultSharePointTenant = async (): Promise<{ hostname: string; url: string }> => {
  const response = await getGraph<Site>('/sites/root?select=webUrl,siteCollection/hostname')

  if (!response.data.siteCollection?.hostname || !response.data.webUrl) {
    throw new Error('Could not resolve default SharePoint tenant')
  }

  return {
    hostname: response.data.siteCollection.hostname,
    url: response.data.webUrl,
  }
}
