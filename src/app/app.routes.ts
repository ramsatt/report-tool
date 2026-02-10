import { Routes } from '@angular/router';
import { DashboardComponent } from './dashboard/dashboard.component';
import { SprintPlanningComponent } from './sprint-planning/sprint-planning.component';
import { SprintClosureComponent } from './sprint-closure/sprint-closure.component';
import { MonthlyReportComponent } from './monthly-report/monthly-report.component';
import { MonthlyConnectDeckComponent } from './monthly-connect-deck/monthly-connect-deck.component';

export const routes: Routes = [
  { path: '', component: DashboardComponent },
  { path: 'sprint-planning', component: SprintPlanningComponent },
  { path: 'sprint-closure', component: SprintClosureComponent },
  { path: 'monthly-report', component: MonthlyReportComponent },
  { path: 'monthly-connect-deck', component: MonthlyConnectDeckComponent },
  // Compatibility with legacy URLs if manually entered
  { path: 'sprint_planning.html', redirectTo: '/sprint-planning', pathMatch: 'full' },
  { path: 'sprint_closure.html', redirectTo: '/sprint-closure', pathMatch: 'full' }
];
