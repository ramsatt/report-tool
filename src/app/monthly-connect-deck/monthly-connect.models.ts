export interface GeneralInfo {
  month: string;
  year: number;
  projectName: string;
  coreHighlights?: string[];
  appHighlights?: string[];
}

export interface DeliveryMetric {
  stream: string;
  sprint: string;
  month: string;
  committed: number;
  delivered: number;
  deliveryPct: number;
  features: string;
  deployedDate: string;
  deploymentStatus: string;
  bugs: number;
  comment: string;
}

export interface VelocityStrategy {
  month: string;
  sprint: string;
  committed: number;
  delivered: number;
  deliveryPct: number;
  plannedDeployment?: string;
  actualDeployment?: string;
  comment: string;
}

export interface DefectItem {
  id: number;
  type: string;
  title: string;
  priority: number;
  createdDate: string;
  resolvedDate: string;
  assignedTo: string;
  iteration: string;
  state: string;
  tags: string;
  defectType?: string;
  resolvedReason?: string;
  rootCause?: string;
}

export interface RoadmapItem {
  sno: number;
  workItem: string;
  details: string;
  type: string;
  priority: string;
  status: string;
}

export interface Feedback {
  date: string;
  actionItem: string;
  owner: string;
  status: string;
  comments: string;
}

export interface LeavePlan {
  date: string;
  event: string;
  member: string;
}

export interface TeamAction {
  actionItem: string;
  duration: string;
  comments: string;
}

export interface MonthlyData {
  generalInfo: GeneralInfo;
  deliveryMetrics: DeliveryMetric[];
  velocityStrategy: VelocityStrategy[];
  defectAnalysis: DefectItem[];
  backlogItems: DefectItem[];
  defectMetrics: DefectItem[];
  roadmap: RoadmapItem[];
  feedback: Feedback[];
  leavePlan: LeavePlan[];
  teamActions: TeamAction[];
}
