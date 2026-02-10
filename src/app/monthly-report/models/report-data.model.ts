/**
 * Data models for Monthly Report slides
 */

export interface DeliveryData {
  sprintMonth: string;
  committed: number;
  delivered: number;
  achieved: string;
  features: string[];
  deploymentStatus: string;
  bugs: number;
  comments?: string;
}

export interface DefectData {
  id: string;
  desc: string;
  priority: string;
  status: string;
  intro?: string;
  created?: string;
  assigned?: string;
  eta?: string;
  statusColor?: string;
  statusText?: string;
}

export interface DefectTrendData {
  id: string;
  desc: string;
  month: string;
  issueType: string;
  closedSprint?: string;
  createdSprint?: string;
}

export interface MigrationData {
  module: string;
  status: string;
  progress: number;
  eta?: string;
}

export interface AutomationData {
  feature: string;
  status: string;
  coverage: number;
  notes?: string;
}

export interface LeavePlanData {
  date: string;
  event: string;
  member: string;
}

export interface TeamActionData {
  item: string;
  duration: string;
  comments: string;
}

export interface SlideConfig {
  currentMonth: string;
  currentYear: string;
}

export interface DeliveryTotals {
  committed: number;
  delivered: number;
  achieved: string;
}
