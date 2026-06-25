export interface IRequest {
  ID: number;
  RequestCode: string;
  Title: string;
  Status: string;
  AuthorTitle: string;
  Created: string;
  AssignedToEmail: string;
  AssignedToTitle: string;
}

export interface IRequestManagementState {
  requests: IRequest[];
  isLoading: boolean;
  error: string | null;
  searchText: string;
  statusFilter: string;
  selectedRequest: IRequest | null;
  isViewPanelOpen: boolean;
  isFormPanelOpen: boolean;
  formTitle: string;
  formRequestCode: string;
  formAssignedToLoginName: string;
  formAssignedToPersonas: { text: string; secondaryText: string; loginName: string }[];
  formKey: number;
  formError: string | null;
  isSubmitting: boolean;
}
