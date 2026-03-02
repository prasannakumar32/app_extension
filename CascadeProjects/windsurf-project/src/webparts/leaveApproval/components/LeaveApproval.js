import * as React from 'react';
import { Pivot, PivotItem, Text } from '@fluentui/react';
import styles from './LeaveApproval.module.scss';
import { ApplyLeaveForm } from './ApplyLeaveForm.js';
import { LeaveHistory } from './LeaveHistory.js';
import { LeaveService } from '../services/LeaveService.js';

export default function LeaveApproval(props) {
  const service = React.useMemo(
    () => new LeaveService(props.webAbsoluteUrl, props.spHttpClient, props.listTitle),
    [props.webAbsoluteUrl, props.spHttpClient, props.listTitle]
  );

  const [refreshToken, setRefreshToken] = React.useState(0);

  return React.createElement(
    'section',
    { className: `${styles.leaveApproval} ${props.hasTeamsContext ? styles.teams : ''}` },
    React.createElement(
      'div',
      { className: styles.container },
      React.createElement(
        'div',
        { className: styles.card },
        React.createElement(
          'div',
          { className: styles.headerRow },
          React.createElement(Text, { variant: 'xLarge' }, 'Leave Management'),
          React.createElement(Text, { className: styles.subtle }, 'Employee view')
        ),
        React.createElement(
          'div',
          { className: styles.pivotArea },
          React.createElement(
            Pivot,
            null,
            React.createElement(
              PivotItem,
              { headerText: 'Apply Leave' },
              React.createElement(ApplyLeaveForm, {
                service,
                onSubmitted: () => setRefreshToken((v) => v + 1)
              })
            ),
            React.createElement(
              PivotItem,
              { headerText: 'Leave History / Status' },
              React.createElement(LeaveHistory, { service, refreshToken })
            )
          )
        )
      )
    )
  );
}
