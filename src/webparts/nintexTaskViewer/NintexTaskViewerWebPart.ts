import {
  Version
} from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import styles from './NintexTaskViewerWebPart.module.scss';
import * as strings from 'NintexTaskViewerWebPartStrings';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';


export interface INintexTaskViewerWebPartProps {
  description: string;
  assignee: string;
  from: string;
  to: string;
  status: string;
  workflowInstanceId: string;
  workflowName: string;
  clientId: string;
  clientSecret: string;
}

interface ITaskAssignment {
  assignee: string;
  completedBy: string;
  completedDate: string;
  status: string;
  urls: {
    formUrl: string;
  };
}

interface ITask {
  name: string;
  workflowName: string;
  status: string;
  createdDate: string;
  assigneeEmail: string;
  completedBy: string;
  dateCompleted: string | undefined;
  openTask: string;
  taskAssignments: ITaskAssignment[];
}


interface IApiResponse {
  tasks: ITask[];
}


export default class NintexTaskViewerWebPart extends BaseClientSideWebPart<INintexTaskViewerWebPartProps> {

  constructor() {
    super();
  }
  
  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      console.log('onInit is called')
      // Check if properties have values, if not set defaults
  
      // Default description to empty string
      this.properties.description = this.properties.description || '';
  
      // Default assignee to current user's email
      this.properties.assignee = this.properties.assignee || this.context.pageContext.user.email;
  
      // Default date range to last 90 days
      const currentDate = new Date();
      this.properties.from = this.properties.from || new Date(currentDate.setDate(currentDate.getDate() - 90)).toISOString();
      this.properties.to = this.properties.to || new Date().toISOString();
  
      // Default status to 'active'
      this.properties.status = this.properties.status || 'active';
  
      this.properties.workflowInstanceId = this.properties.workflowInstanceId || '';
      this.properties.workflowName = this.properties.workflowName || '';
      this.properties.clientId = this.properties.clientId || '';
      this.properties.clientSecret = this.properties.clientSecret || '';
    });
  }
  
  
  private safelyAttachEventListener(selector: string, event: string, handler: EventListenerOrEventListenerObject): void {
    const element = this.domElement.querySelector(selector);
    if (element) {
        element.addEventListener(event, handler);
    }
  }

  private handleAssigneeChange(event: Event): void {
    const selectedValue = (event.target as HTMLSelectElement).value;
    this.properties.assignee = selectedValue === 'MyTasks' ? this.context.pageContext.user.email : '';
  }
  
  private handleDateRangeChange(event: Event): void {
    const selectedValue = (event.target as HTMLSelectElement).value;
    const currentDate = new Date();
  
    switch(selectedValue) {
      case "last90":
        this.properties.from = new Date(currentDate.setDate(currentDate.getDate() - 90)).toISOString();
        this.properties.to = new Date().toISOString();
        break;
      case "last180":
        this.properties.from = new Date(currentDate.setDate(currentDate.getDate() - 180)).toISOString();
        this.properties.to = new Date().toISOString();
        break;
      case "thisYear":
        this.properties.from = new Date(currentDate.getFullYear(), 0, 1).toISOString();
        this.properties.to = new Date(currentDate.getFullYear(), 11, 31).toISOString();
        break;
      case "thisAndLastYear":
        this.properties.from = new Date(currentDate.getFullYear() - 1, 0, 1).toISOString();
        this.properties.to = new Date(currentDate.getFullYear(), 11, 31).toISOString();
        break;
    }
  }
  
  private handleStatusChange(event: Event): void {
    this.properties.status = (event.target as HTMLSelectElement).value;
  }
  
  private handleWorkflowInstanceIdChange(event: Event): void {
    this.properties.workflowInstanceId = (event.target as HTMLInputElement).value;
  }
  
  private handleWorkflowNameChange(event: Event): void {
    this.properties.workflowName = (event.target as HTMLInputElement).value;
  }
  
  private fetchTasks(): void {
    this.getTasks()
      .then((tasks: ITask[]) => {
        const tasksHtml = this.generateTaskHTML(tasks);
        const tasksOutputElement = this.domElement.querySelector('#tasksOutput');
        if (tasksOutputElement) {
          tasksOutputElement.innerHTML = tasksHtml;
        } else {
          console.warn('Element with ID #tasksOutput not found in the DOM.');
        }
      })
      .catch((error: unknown) => {
        console.error(error);
        const errorMessage = (error instanceof Error) ? error.message : 'An error occurred';
        this.domElement.innerHTML = `
          <section class="${styles.nintexTaskViewer}">
            <h2>Error loading tasks</h2>
            <p>${escape(errorMessage)}</p>
          </section>
        `;
      });
  }
  

  private generateTaskHTML(tasks: ITask[]): string {
    return `
        <table>
            <thead>
                <tr>
                    <th>Task Name</th>
                    <th>Workflow Name</th>
                    <th>Status</th>
                    <th>Date Initiated</th>
                    <th>Assignee Email</th>
                    <th>Completed by</th>
                    <th>Date Completed</th>
                    <th>Open Task</th>
                </tr>
            </thead>
            <tbody>
                ${tasks.map(task => `
                <tr>
                    <td>${task.name}</td>
                    <td>${task.workflowName}</td>
                    <td>${task.status}</td>
                    <td>${new Date(task.createdDate).toLocaleDateString()}</td>
                    <td>${task.assigneeEmail}</td>
                    <td>${task.completedBy}</td>
                    <td>${task.dateCompleted ? new Date(task.dateCompleted).toLocaleDateString() : ''}</td>
                    <td><a href="${task.openTask}" target="_blank">Open</a></td>
                </tr>
                `).join('')}
            </tbody>
        </table>
    `;
  }

  public render(): void {
    console.log("Status property before render:", this.properties.status);
    this.domElement.innerHTML = `
    <section class="${styles.nintexTaskViewer}">
      <h2>Task Filters</h2>
      <div class="filterBar">
        <div>
          Task View: 
          <select id="assigneeDropdown">
            <option selected value="MyTasks">My Tasks</option>
            <option value="AllTasks">All Tasks</option>
          </select>
        </div>
        <div>
          Date Range: 
          <select id="dateRangeDropdown">
            <option selected value="last90">Last 90 days</option>
            <option value="last180">Last 180 days</option>
            <option value="thisYear">This calendar year</option>
            <option value="thisAndLastYear">This year and last year</option>
          </select>
        </div>
        <div>
          Status: 
          <select id="statusDropdown">
            <option selected value="active">Active</option>
            <option value="expired">Expired</option>
            <option value="complete">Complete</option>
            <option value="overridden">Overridden</option>
            <option value="terminated">Terminated</option>
            <option value="all">All</option>
          </select>
        </div>
        <div>
          Workflow Instance ID: <input type="text" id="workflowInstanceIdInput">
          Workflow Name: <input type="text" id="workflowNameInput">
        </div>
        <button id="fetchTasksBtn">Fetch Tasks</button>
      </div>
      <h2>Tasks</h2>
      <div id="tasksOutput"></div>
    </section>
  `;
  
  

  this.safelyAttachEventListener('#assigneeDropdown', 'change', this.handleAssigneeChange.bind(this));
  this.safelyAttachEventListener('#dateRangeDropdown', 'change', this.handleDateRangeChange.bind(this));
  this.safelyAttachEventListener('#statusDropdown', 'change', this.handleStatusChange.bind(this));
  this.safelyAttachEventListener('#workflowInstanceIdInput', 'input', this.handleWorkflowInstanceIdChange.bind(this));
  this.safelyAttachEventListener('#workflowNameInput', 'input', this.handleWorkflowNameChange.bind(this));
  this.safelyAttachEventListener('#fetchTasksBtn', 'click', this.fetchTasks.bind(this));

  this.fetchTasks();
  }

  //NintexTaskListApp

protected getTasks(): Promise<ITask[]> {
  // Start with the base URL
  let url = 'https://eu.nintex.io/workflows/v2/tasks?';

  // Use the properties set by the event handlers
  const {
    workflowName,
    assignee,
    from,
    to,
    workflowInstanceId,
    status: rawStatus
  } = this.properties;

  const status = rawStatus === 'all' ? null : rawStatus;

  // Construct the query parameters
  const queryParams = {
    workflowName,
    assignee,
    from,
    to,
    status,
    workflowInstanceId
  };

  const queryString = Object.entries(queryParams)
    .filter(([, value]) => value)  // Remove empty values
    .map(([key, value]) => `${key}=${encodeURIComponent(value as string)}`)
    .join('&');

  url += queryString;

  console.log(`Generated Request URL: ${url}`);
  console.log('Properties before URL generation:', this.properties);

  // Call getAccessToken to fetch the access token
  return this.getAccessToken().then((accessToken: string) => {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Authorization', `Bearer ${accessToken}`);
    requestHeaders.append('Accept', 'application/json');

    const httpClientOptions: IHttpClientOptions = {
      headers: requestHeaders
    };

    return this.context.httpClient.get(url, HttpClient.configurations.v1, httpClientOptions);
  })
  .then((res: HttpClientResponse) => {
    if (!res.ok) {
        throw new Error(`API returned status: ${res.status}`);
    }
    return res.json();
  })
  .then((data: IApiResponse): ITask[] => { 
    if (!Array.isArray(data.tasks)) {
        throw new Error('Tasks is not an array or is missing from API response');
    }
    return data.tasks.map((task: ITask): ITask => ({
        name: task.name,
        workflowName: task.workflowName,
        status: task.status,
        createdDate: task.createdDate,
        assigneeEmail: task.taskAssignments[0]?.assignee || '',
        completedBy: task.taskAssignments[0]?.completedBy || '',
        dateCompleted: task.taskAssignments[0]?.completedDate || '',
        openTask: task.taskAssignments[0]?.urls?.formUrl || '',
        taskAssignments: task.taskAssignments
    }));
  });
}

private getAccessToken(): Promise<string> {
  const clientId = this.properties.clientId;
  const clientSecret = this.properties.clientSecret;

  const url = 'https://eu.nintex.io/authentication/v1/token';
  const body = {
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: 'client_credentials'
  };
  const headers: Headers = new Headers();
  headers.append('Content-Type', 'application/json');

  return this.context.httpClient.post(url, HttpClient.configurations.v1, {
      body: JSON.stringify(body),
      headers: headers
  })
  .then((res: HttpClientResponse) => {
      if (!res.ok) {
          throw new Error(`Server responded with status: ${res.status}`);
      }
      return res.json();
  })
  .then((data: { access_token: string }) => {
      if (!data.access_token) {
          throw new Error("No access token found in response");
      }
      return data.access_token;
  })
  .catch((err) => {
      console.error("Error in getAccessToken:", err);
      throw err;  // Re-throw the error to handle it outside
  });
}


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [

        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('clientId', {
                  label: "Client ID",
                  value: this.properties.clientId
                }),
                PropertyPaneTextField('clientSecret', {
                  label: "Client Secret",
                  value: this.properties.clientSecret
                }),                
              ]
            }
          ]
        }
      ]
    };
  }
}