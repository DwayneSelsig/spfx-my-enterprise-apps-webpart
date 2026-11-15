/** Optional initial filter; searching stays inside the existing app picker. */
export interface IEnterpriseAppsCopilotComponentProperties {
  query?: string;
}

export default {
  type: 'object',
  properties: {
    query: {
      type: 'string',
      description: 'App name or part of an app name to find. Omit or use an empty string to show all apps.'
    }
  },
  additionalProperties: false
};
