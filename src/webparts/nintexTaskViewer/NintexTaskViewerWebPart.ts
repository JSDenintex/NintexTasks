import {
  Version
} from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import styles from './NintexTaskViewerWebPart.module.scss';
import * as strings from 'NintexTaskViewerWebPartStrings';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';


type TaskStatus = 'active' | 'expired' | 'complete' | 'overridden' | 'terminated';

export interface INintexTaskViewerWebPartProps {
  description: string;
  assignee: string;
  from: string;
  status: string;
  workflowId: string;
  workflowName: string;
  boolAllTasks: boolean;
  clientId: string;
  clientSecret: string;
  assigneeEmail: string;
  tenancyRegion: string;
  autoRefreshInterval: string;
  enableAutoRefresh: boolean;
  authMethod: string;
}

interface ITaskAssignment {
  id: string;
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
  id: string;
  workflowName: string;
  status: string;
  createdDate: string;
  assigneeEmail: string;
  completedBy: string;
  dateCompleted: string | undefined;
  openTask: string;
  taskAssignments: ITaskAssignment[];
  outcomes: string[] | undefined;
  message: string;
}


interface IApiResponse {
  tasks: ITask[];
}

export default class NintexTaskViewerWebPart extends BaseClientSideWebPart<INintexTaskViewerWebPartProps> {

  constructor() {
    super();
  }

  private fetchedTasks: ITask[] = [];

  //API FUNCTIONS

  private getApiEndpoint(): string {
    switch (this.properties.tenancyRegion) {
      case 'us':
        return 'https://us.nintex.io';
      case 'eu':
        return 'https://eu.nintex.io';
      case 'au':
        return 'https://au.nintex.io';
      case 'ca':
        return 'https://ca.nintex.io';
      case 'uk':
        return 'https://uk.nintex.io';
      default:
        return 'https://us.nintex.io';
    }
  }

  private fetchTasks(makeAPICall: boolean = true): void {
    if (makeAPICall) {
      this.getTasks()
        .then((tasks: ITask[]) => {
          this.fetchedTasks = tasks; // Store the fetched tasks
          this.renderSortedTasks(); // Call renderSortedTasks to render the tasks
        })
        .catch((error: unknown) => {
          console.error(error);
          const errorMessage = (error instanceof Error) ? error.message : 'An error occurred';
          this.domElement.innerHTML = `
            <section class="${styles.nintexTaskViewer}">
              <h2>Unable to load tasks</h2>
              <p>${escape(errorMessage)}</p>
              <p>Please ensure that the webpart properties are configured correctly.</p>
            </section>
          `;
        });
    }
  }

  protected async getTasks(): Promise<ITask[]> {
    // Start with the base URL
    const baseUrl = this.getApiEndpoint();
    let url = `${baseUrl}/workflows/v2/tasks?`;
  
    // Use the properties set by the event handlers
    const {
      workflowName,
      assignee,
      from,
      status: rawStatus
    } = this.properties;
  
    const status = rawStatus;
  
    // Construct the query parameters
    const queryParams: {
      workflowName?: string;
      assignee?: string;
      from?: string;
      status?: string;
    } = {
      workflowName,
      assignee,
      from,
      status,
    };
  
    // Filter out null, undefined, and empty string values
    const filteredQueryParams: { [key: string]: string } = {};
    for (const [key, value] of Object.entries(queryParams)) {
      if (value !== null && value !== undefined && value !== '') {
        filteredQueryParams[key] = value;
      }
    }
  
    const queryString = Object.entries(filteredQueryParams)
      .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
      .join('&');
  
    // Append the unique cache-busting parameter
    const cacheBuster = `cacheBuster=${new Date().getTime()}`;
    url += queryString + (queryString ? '&' : '') + cacheBuster;
  
    console.log(`Generated Request URL: ${url}`);
    // console.log('Properties before URL generation:', this.properties);
  
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
        id: task.id,
        status: task.status,
        createdDate: task.createdDate,
        assigneeEmail: task.taskAssignments[0]?.assignee || '',
        completedBy: task.taskAssignments[0]?.completedBy || '',
        dateCompleted: task.taskAssignments[0]?.completedDate || '',
        openTask: task.taskAssignments[0]?.urls?.formUrl || '',
        taskAssignments: task.taskAssignments,
        outcomes: task.outcomes || undefined,
        message: task.message,
      }));
    });
  }  
  
  private async getAccessToken(): Promise<string> {
    const clientId = this.properties.clientId;
    const clientSecret = this.properties.clientSecret;
    const baseUrl = this.getApiEndpoint();
    const url = `${baseUrl}/authentication/v1/token`;
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

  private confirmTaskCompletion(taskId: string, outcome: string, taskName: string, assignmentId: string): void {
    const endpoint = this.getApiEndpoint();
    this.getAccessToken().then(accessToken => {
      const url = `${endpoint}/workflows/v2/tasks/${taskId}/assignments/${assignmentId}`;
      const options = {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json, application/problem+json',
          Authorization: `Bearer ${accessToken}` // Retrieve the access token using the method
        },
        body: JSON.stringify({ outcome: outcome })
      };
  
      fetch(url, options)
        .then(response => response.json())
        .then(data => {
          console.log('Task completed:', data);
          // Close the modal and possibly update the UI
          const modal = document.getElementById('confirmationModal');
          if (modal) {
            modal.style.display = 'none';
            this.fetchTasks();
          }
          // Additional logic to update UI
        })
        .catch(error => console.error('Error completing task:', error));
    }).catch(error => {
      console.error('Error getting access token:', error);
    });
  }

  //END API CALLS
  //UTILITY
  
  protected async onInit(): Promise<void> {
    return super.onInit().then(() => {
      console.log('onInit is called')
      // Check if properties have values, if not set defaults
      this.properties.enableAutoRefresh = this.properties.enableAutoRefresh || false;
  
      // Default description to empty string
      this.properties.description = this.properties.description || '';

      // Default date range to last 90 days
      const currentDate = new Date();
      this.properties.from = this.properties.from || new Date(currentDate.setDate(currentDate.getDate() - 90)).toISOString();
  
      // Default status to 'active'
      this.properties.status = 'active';
  
      // Check if assigneeEmail property is filled
      if (this.properties.assigneeEmail) {
        console.log("Using assigneeEmail input value:", this.properties.assigneeEmail);
        this.properties.assignee = this.properties.assigneeEmail;
      } else {
        // Check the viewDropdown selection
        const viewDropdown = document.getElementById('viewDropdown') as HTMLSelectElement;
        if (viewDropdown && viewDropdown.value === 'AllTasks') {
          // If All Tasks selected, remove assignee from the API call
          this.properties.assignee = '';
        } else {
          console.log("Using default/context user email:", this.context.pageContext.user.email);
          this.properties.assignee = this.context.pageContext.user.email;
        }
      }
      this.fetchTasks();
      this.startAutoRefresh();
    });
  }
  
  private safelyAttachEventListener(selector: string, event: string, handler: EventListenerOrEventListenerObject): void {
    const element = this.domElement.querySelector(selector);
    if (element) {
        element.addEventListener(event, handler);
    }
  }

  private handleViewChange(event: Event): void {
    const selectedValue = (event.target as HTMLSelectElement).value;
    if (selectedValue === 'MyTasks') {
      this.properties.assignee = this.context.pageContext.user.email;
    } else {
      this.properties.assignee = '';
    }
    // Call fetchTasks() after setting the properties
    this.fetchTasks();
  }
  
  private handleDateRangeChange(event: Event): void {
    const selectedValue = (event.target as HTMLSelectElement).value;
    const currentDate = new Date();
  
    // Function to format a date to YYYY-MM-DD
    const formatDate = (date: Date): string => {
      return date.toISOString().split('T')[0];
    };
  
    switch(selectedValue) {
      case "last90":
        this.properties.from = formatDate(new Date(currentDate.setDate(currentDate.getDate() - 90)));
        break;
      case "last180":
        this.properties.from = formatDate(new Date(currentDate.setDate(currentDate.getDate() - 180)));
        break;
      case "thisYear":
        this.properties.from = formatDate(new Date(currentDate.getFullYear(), 0, 1));
        break;
      case "thisAndLastYear":
        this.properties.from = formatDate(new Date(currentDate.getFullYear() - 1, 0, 1));
        break;
    }
    // Call fetchTasks() after setting the properties
    this.fetchTasks();
  }
  
    
  private handleStatusChange(event: Event): void {
    this.properties.status = (event.target as HTMLSelectElement).value;
    console.log("Status Changed - setting status to ", this.properties.status)
    // Call fetchTasks() after setting the properties
    this.fetchTasks();
  }
  
  private handleButtonClick(): void {
    // Set shouldMakeAPICall to true when the button is clicked
    this.fetchTasks();
  }

  private setUpEventDelegation(): void {
    const container = this.domElement.querySelector('#NintexTaskViewerWP');
    console.log("Setting up event delegation")
    console.log(container)
    if (container) {
      container.addEventListener('click', (event) => {
        const target = event.target as HTMLElement;
        if (target && target.hasAttribute('data-outcome')) {
          // Existing logic to handle the button click
          const taskId = target.getAttribute('data-task-id');
  
          if (taskId) {
            const task = this.findTaskById(taskId);
            if (task) {
              this.handleOutcomeButtonClick(event, task);
            }
          } else {
            console.warn('Task ID not found on clicked element');
          }
        }
      });
    }
  }
  
  private findTaskById(taskId: string): ITask | undefined {
    return this.fetchedTasks.find(task => task.id === taskId);
  }
  
  private handleOutcomeButtonClick(event: Event, task: ITask): void {
    // Handle the outcome button click here
    const target = event.target as HTMLElement;
    console.log(`Outcome button clicked with data-outcome: ${target.getAttribute("data-outcome")}`);
    const taskId = target.getAttribute("data-task-id");
    const outcome = target.getAttribute("data-outcome");
    const taskName = target.getAttribute("data-task-name");
    const assignmentId = task.taskAssignments[0].id;
  
    if (taskId && outcome && taskName) {
      this.showConfirmationModal(taskId, outcome, taskName, assignmentId);
    }
  }
  
  private showConfirmationModal(taskId: string, outcome: string, taskName: string, assignmentId: string): void {
    // Check if an existing modal exists and remove it
    const existingModal = document.getElementById('confirmationModal');
    if (existingModal) {
      existingModal.remove();
    }
    
    // Create modal elements
    const modal = document.createElement('div');
    modal.id = 'confirmationModal';
    modal.classList.add(styles.confirmationModal); // Use styles.[class]
  
    const modalContent = document.createElement('div');
    modalContent.classList.add(styles.modalContent); // Use styles.[class]
  
    const message = document.createElement('p');
    message.innerHTML = `Are you sure you want to complete the task with the outcome of <strong>${outcome}</strong>?`;
  
    const confirmButton = document.createElement('button');
    confirmButton.id = 'confirmButton';
    confirmButton.classList.add(styles.confirmButton); // Use styles.[class]
    confirmButton.textContent = 'Confirm';
  
    const cancelButton = document.createElement('button');
    cancelButton.id = 'cancelButton';
    cancelButton.classList.add(styles.cancelButton); // Use styles.[class]
    cancelButton.textContent = 'Cancel';
  
    // Append elements to the modal
    modalContent.appendChild(message);
    modalContent.appendChild(confirmButton);
    modalContent.appendChild(cancelButton);
    modal.appendChild(modalContent);
  
    // Append the modal to the DOM
    document.body.appendChild(modal);
  
  // Attach event listeners with console logging for debugging
  confirmButton.onclick = () => {
    console.log(`Confirm button clicked for Task ID: ${taskId}, Outcome: ${outcome}`);
    this.confirmTaskCompletion(taskId, outcome, taskName, assignmentId);
  };

  cancelButton.onclick = () => {
    console.log('Cancel button clicked');
    modal.style.display = 'none';
  };
  
    // Display the modal
    modal.style.display = 'block';
  }

  // toggle

  private readonly SortDownIcon: string = `
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi bi-sort-down-alt" viewBox="0 0 16 16">
      <path d="M3.5 3.5a.5.5 0 0 0-1 0v8.793l-1.146-1.147a.5.5 0 0 0-.708.708l2 1.999.007.007a.497.497 0 0 0 .7-.006l2-2a.5.5 0 0 0-.707-.708L3.5 12.293zm4 .5a.5.5 0 0 1 0-1h1a.5.5 0 0 1 0 1zm0 3a.5.5 0 0 1 0-1h3a.5.5 0 0 1 0 1zm0 3a.5.5 0 0 1 0-1h5a.5.5 0 0 1 0 1zM7 12.5a.5.5 0 0 0 .5.5h7a.5.5 0 0 0 0-1h-7a.5.5 0 0 0-.5.5"/>
    </svg>
    `;

  private readonly SortUpIcon: string = `
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi bi-sort-up" viewBox="0 0 16 16">
      <path d="M3.5 12.5a.5.5 0 0 1-1 0V3.707L1.354 4.854a.5.5 0 1 1-.708-.708l2-1.999.007-.007a.498.498 0 0 1 .7.006l2 2a.5.5 0 1 1-.707.708L3.5 3.707zm3.5-9a.5.5 0 0 1 .5-.5h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1-.5-.5M7.5 6a.5.5 0 0 0 0 1h5a.5.5 0 0 0 0-1zm0 3a.5.5 0 0 0 0 1h3a.5.5 0 0 0 0-1zm0 3a.5.5 0 0 0 0 1h1a.5.5 0 0 0 0-1z"/>
    </svg>
  `;

  private sortAscending: boolean = true;
  
  private handleSortToggle(): void {
    this.sortAscending = !this.sortAscending;
    this.updateSortIcon(); 
    this.renderSortedTasks();
  }
  
  private renderSortedTasks(): void {
    // Sort the fetched tasks
    const sortedTasks = this.fetchedTasks.sort((a, b) => {
      return this.sortAscending
        ? new Date(a.createdDate).getTime() - new Date(b.createdDate).getTime()
        : new Date(b.createdDate).getTime() - new Date(a.createdDate).getTime();
    });
  
    // Generate HTML for sorted tasks
    const tasksHtml = this.generateTaskHTML(sortedTasks);
  
    // Render the sorted tasks
    const tasksOutputElement = this.domElement.querySelector('#tasksOutput');
    if (tasksOutputElement) {
      tasksOutputElement.innerHTML = tasksHtml;
    } else {
      console.warn('Element with ID #tasksOutput not found in the DOM.');
    }
  }

  private updateSortIcon(): void {
    const sortIconElement = this.domElement.querySelector('#sortIcon');
    if (sortIconElement) {
      const sortIconHtml = this.sortAscending ? this.SortUpIcon : this.SortDownIcon;
      sortIconElement.innerHTML = sortIconHtml;
    }
  }

  //RENDER COMPONENTS

  private generateTaskHTML(tasks: ITask[]): string {
    if (tasks.length === 0) {
      return `
        <div class="${styles.taskItem}">
          <div class="${styles.noTasksFound}">No tasks found</div>
        </div>
      `;
    }

    function getStatusClass(status: string): string {
      if (isTaskStatus(status)) {
        return styles[status];
      }
      console.warn(`Unexpected status: ${status}`);
      return '';
    }

    function toCapitalCase(str: string): string {
      return str.split(' ')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(' ');
    }
    

    // Helper type guard to check if a string is a valid task status
    function isTaskStatus(status: string): status is TaskStatus {
      return ['active', 'expired', 'complete', 'overridden', 'terminated'].includes(status);
    }

    return `
    
    <div class="${styles.taskItems}">
      ${tasks.map(task => `
        <div class="${styles.taskItem}">
          <div class="${styles.taskHeader}">
          <span class="${styles.statusSpan} ${getStatusClass(task.status as TaskStatus)}">${toCapitalCase(task.status)}</span>
            ${this.renderOutcomeButtons(task)}
          </div>
            <div class="${styles.taskBody}">
              <div class="${styles.taskDetails}">
                <div class="${styles.taskName}">${task.name}</div>
                <div class="${styles.workflowName}">${task.workflowName}</div>
              </div>
              <div class="${styles.taskDetails}">
                <div class="${styles.assigneeEmail}">Assignee Email: ${task.assigneeEmail}</div>
                <div class="${styles.dateInitiated}">Date Initiated: ${new Date(task.createdDate).toLocaleDateString()}</div>
              </div>
              ${task.status !== 'active' ? `
              <div class="${styles.taskDetails}">
                <div class="${styles.completedBy}">Completed by: ${task.completedBy}</div>
                <div class="${styles.dateCompleted}">Date Completed: ${task.dateCompleted ? new Date(task.dateCompleted).toLocaleDateString() : ''}</div>
              </div>
            ` : ''}
            </div>
            <div class="${styles.taskBody}">
            <div class="${styles.taskDetails}">
              <div class="${styles.taskName}">${task.message}</div>
            </div>
          </div>
          </div>
        `).join('')}
      </div>
    `;
    
  }

  private renderOutcomeButtons(task: ITask): string {
    if (task.openTask && task.status === 'active') {
      // If a form URL (openTask) is present, render the open button
      return `
        <a class="${styles.openButton}" href="${task.openTask}" target="_blank">Open</a>
      `;
    } else if (task.status === 'active' && task.outcomes && task.outcomes.length > 0) {
      // If no form URL and outcomes are present, render outcome buttons
      const outcomeButtons = task.outcomes.map(outcome => `
        <button
          class="${styles.outcomeButton}"
          id="outcomeButton-${task.id}-${outcome}"
          data-task-id="${task.id}"
          data-outcome="${outcome}"
          data-task-name="${task.name}"
        >
          ${outcome}
        </button>
        
      `);
      return outcomeButtons.join('');
    } else {
      // If neither form URL nor outcomes are present or task is not active, return an empty string
      return '';
    }
  }

  public render(): void {
    // Generate the HTML for the dropdown based on whether assigneeEmail is filled
    const viewDropdownHtml = this.properties.boolAllTasks ? '' : `
    <div class="${styles.filterItem}">
      <label for="viewDropdown">Task View:</label>
      <select class="${styles.filterDropdown}" id="viewDropdown">
        <option selected value="MyTasks">My Tasks</option>
        <option value="AllTasks">All Tasks</option>
      </select>
    </div>
  `;
  
    this.domElement.innerHTML = `
      <div id="NintexTaskViewerWP">
        <section class="${styles.nintexTaskViewer}">
          <div class="${styles.filterGroup}">
            <div class="${styles.filterBar}">
              ${viewDropdownHtml}
              <div class="${styles.filterItem}">
                <label for="dateRangeDropdown">Date Range:</label>
                <select class="${styles.filterDropdown}" id="dateRangeDropdown">
                  <option selected value="last90">Last 90 days</option>
                  <option value="last180">Last 180 days</option>
                  <option value="thisYear">This calendar year</option>
                  <option value="thisAndLastYear">This year and last year</option>
                </select>
              </div>
              <div class="${styles.filterItem}">
                <label for="statusDropdown">Status:</label>
                <select class="${styles.filterDropdown}" id="statusDropdown">
                  <option selected value="active">Active</option>
                  <option value="expired">Expired</option>
                  <option value="complete">Complete</option>
                  <option value="overridden">Overridden</option>
                  <option value="terminated">Terminated</option>
                  <option value="all">All</option>
                </select>
              </div>
            </div>
          </div>
          <div class="${styles.tasksContainer}">
            <h2>Tasks</h2>
            <button id="sortTasksBtn"><span id="sortIcon">${this.sortAscending ? this.SortUpIcon : this.SortDownIcon}</span> Sort by Date</button>
            <button id="fetchTasksBtn">Update</button>
          </div>
          <div id="tasksOutput"></div>
        </section>
      </div>
    `;

    // Attach event listener for the sorting button
    const sortTasksBtn = this.domElement.querySelector('#sortTasksBtn');
    if (sortTasksBtn) {
      sortTasksBtn.addEventListener('click', this.handleSortToggle.bind(this));
    }
  
    // Attach event listeners and fetch tasks as usual
    this.safelyAttachEventListener('#viewDropdown', 'change', this.handleViewChange.bind(this));
    this.safelyAttachEventListener('#dateRangeDropdown', 'change', this.handleDateRangeChange.bind(this));
    this.safelyAttachEventListener('#statusDropdown', 'change', this.handleStatusChange.bind(this));
    this.setUpEventDelegation();

    // Attach event listener for button click
    const fetchTasksBtn = this.domElement.querySelector('#fetchTasksBtn');
    if (fetchTasksBtn) {
      fetchTasksBtn.addEventListener('click', this.handleButtonClick.bind(this));
    }
  }

  //END RENDER COMPONENTS
  //PROPERTY PANE CONFIG

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //AUTO REFRESH

  private intervalId: number | undefined;

  protected onDispose(): void {
    this.stopAutoRefresh();
  }

  private startAutoRefresh(): void {
    const intervalSeconds = Number(this.properties.autoRefreshInterval);
  
    // Check if interval is a valid number and greater than 0
    if (!isNaN(intervalSeconds) && intervalSeconds > 0) {
      const intervalMilliseconds = intervalSeconds * 1000;
      this.intervalId = window.setInterval(() => this.fetchTasks(), intervalMilliseconds);
    }
  }
  

  private stopAutoRefresh(): void {
    if (this.intervalId) {
      window.clearInterval(this.intervalId);
    }
  }

  //END AUTO REFRESH

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [

        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Authentication",
              groupFields: [
                PropertyPaneTextField('clientId', {
                  label: "Client ID",
                  value: this.properties.clientId
                }),
                PropertyPaneTextField('clientSecret', {
                  label: "Client Secret",
                  value: this.properties.clientSecret,
                }),
                PropertyPaneDropdown('tenancyRegion', {
                  label: "Tenancy Region",
                  options: [
                    { key: 'us', text: 'West US' },
                    { key: 'eu', text: 'North Europe' },
                    { key: 'au', text: 'Australia' },
                    { key: 'ca', text: 'Canada' },
                    { key: 'uk', text: 'United Kingdom' }
                  ]
                })    
              ]
            },
            {
              groupName: "Filter options",
              groupFields: [
                PropertyPaneTextField('workflowName', {
                  label: 'Filter by Workflow Name',
                }),
                PropertyPaneTextField('assigneeEmail', {
                  label: 'Filter by target user',
                }),
              ]
            },
            {
              groupName: "Auto Refresh Settings",
              groupFields: [
                PropertyPaneToggle('enableAutoRefresh', {
                  label: "Enable Auto Refresh",
                  onText: "Enabled",
                  offText: "Disabled"
                }),
                ...this.properties.enableAutoRefresh ? [
                  PropertyPaneSlider('autoRefreshInterval', {
                    label: 'Auto Refresh Interval (in seconds)',
                    min: 15,
                    max: 300,
                    step: 15,
                    showValue: true,
                    value: 15
                  })
                ] : []
              ]
            }
          ]
        }
      ]
    };
  }

  private propertyChanged: boolean = false;

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
    if (propertyPath === 'assigneeEmail') {
      // Existing logic for assigneeEmail
      const isAssigneeEmailFilled = !!newValue;
      this.properties.assignee = isAssigneeEmailFilled ? newValue : this.context.pageContext.user.email;
      this.render();
    } 
    else if (propertyPath === 'autoRefreshInterval') {
      // Logic for autoRefreshInterval
      const newIntervalSeconds = Number(newValue);
  
      // Restart auto-refresh if the new value is a valid number and greater than 0
      // Stop auto-refresh if the new value is invalid
      if (!isNaN(newIntervalSeconds) && newIntervalSeconds > 0) {
        this.stopAutoRefresh();
        this.startAutoRefresh();
      } else {
        this.stopAutoRefresh();
      }
    }
    this.propertyChanged = true;
  }
  
  protected onPropertyPaneConfigurationComplete(): void {
    // Check the flag
    if (this.propertyChanged) {
      this.fetchTasks();
      // Reset the flag
      this.propertyChanged = false;
    }
  }
  
}
