import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-slide-7-migration',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './slide-7-migration.component.html',
  styleUrls: ['./slide-7-migration.component.css']
})
export class Slide7MigrationComponent {
  // Static data for migration status
  stats = {
    totalCompletion: '100%',
    deployedDate: '31 JAN 2026',
    clusterEnv: 'PRODUCTION',
    overallStatus: 'COMPLETED'

  };


  modules = [
    {
      name: 'Home Dashboard',
      lifecycle: '19 Mar — 01 Apr',
      progress: '100%',
      health: 'COMPLETED',
      notes: 'Demo Verified: DEV',
      isCompleted: true
    },
    {
      name: 'User Specific Screen',
      lifecycle: '19 Mar — 13 May',
      progress: '100%',
      health: 'COMPLETED',
      notes: 'Demo Verified: DEV',
      isCompleted: true
    },
    {
      name: 'Custom Field Logic',
      lifecycle: '30 Apr — 27 May',
      progress: '100%',
      health: 'COMPLETED',
      notes: 'Demo Verified: DEV',
      isCompleted: true
    },
    {
      name: 'Fleet List System',
      lifecycle: '13 Oct — 28 Oct',
      progress: '100%',
      health: 'COMPLETED',
      notes: 'Demo Verified: INT',
      isCompleted: true
    },
    {
      name: 'Digital Factory Core',
      lifecycle: '19 Aug — 16 Sep',
      progress: '100%',
      health: 'COMPLETED',
      notes: 'Demo Verified: INT',
      isCompleted: true
    },
    {
      name: 'Report Group Logic',
      lifecycle: '12 Nov — 25 Nov',
      progress: '100%',
      health: 'COMPLETED',
      notes: 'Demo Verified: INT',
      isCompleted: true
    }
  ];

}
