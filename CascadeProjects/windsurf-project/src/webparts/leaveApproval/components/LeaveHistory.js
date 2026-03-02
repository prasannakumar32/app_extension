import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  MessageBar,
  MessageBarType,
  Spinner,
  Stack,
  Dropdown
} from '@fluentui/react';

const statusOptions = [
  { key: 'All', text: 'All' },
  { key: 'Pending', text: 'Pending' },
  { key: 'Approved', text: 'Approved' },
  { key: 'Rejected', text: 'Rejected' }
];

function formatDate(d) {
  try {
    return new Date(d).toLocaleDateString();
  } catch {
    return d;
  }
}

export function LeaveHistory(props) {
  const [items, setItems] = React.useState([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState('All');

  const load = React.useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      const data = await props.service.getMyLeaves();
      setItems(data);
    } catch (e) {
      setError(e && e.message ? e.message : 'Failed to load leave history.');
    } finally {
      setLoading(false);
    }
  }, [props.service]);

  React.useEffect(() => {
    load().catch(() => {
      // handled
    });
  }, [load, props.refreshToken]);

  const filteredItems = React.useMemo(() => {
    if (statusFilter === 'All') return items;
    return items.filter((i) => i.Status === statusFilter);
  }, [items, statusFilter]);

  const columns = [
    { key: 'id', name: 'ID', fieldName: 'Id', minWidth: 40, maxWidth: 60 },
    { key: 'type', name: 'Type', fieldName: 'LeaveType', minWidth: 90, maxWidth: 140 },
    { key: 'start', name: 'Start', minWidth: 90, onRender: (item) => formatDate(item.StartDate) },
    { key: 'end', name: 'End', minWidth: 90, onRender: (item) => formatDate(item.EndDate) },
    { key: 'half', name: 'Half Day', minWidth: 70, onRender: (item) => (item.IsHalfDay ? 'Yes' : 'No') },
    { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 80, maxWidth: 100 },
    { key: 'comments', name: 'Comments', minWidth: 120, onRender: (item) => item.ApproverComments || '' },
    { key: 'created', name: 'Applied On', minWidth: 100, onRender: (item) => formatDate(item.Created) }
  ];

  return React.createElement(
    Stack,
    { tokens: { childrenGap: 12 } },
    error ? React.createElement(MessageBar, { messageBarType: MessageBarType.error }, error) : null,
    React.createElement(Dropdown, {
      label: 'Filter by Status',
      options: statusOptions,
      selectedKey: statusFilter,
      onChange: (_, opt) => setStatusFilter(String(opt && opt.key))
    }),
    loading ? React.createElement(Spinner, null) : null,
    React.createElement(DetailsList, {
      items: filteredItems,
      columns,
      layoutMode: DetailsListLayoutMode.justified,
      isHeaderVisible: true
    })
  );
}
