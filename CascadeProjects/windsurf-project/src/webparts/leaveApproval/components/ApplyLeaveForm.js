import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  Dropdown,
  TextField,
  DatePicker,
  DayOfWeek,
  Toggle,
  MessageBar,
  MessageBarType,
  Stack,
  Spinner,
  SpinnerSize
} from '@fluentui/react';

const leaveTypeOptions = [
  { key: 'Casual', text: 'Casual' },
  { key: 'Sick', text: 'Sick' },
  { key: 'Earned', text: 'Earned' },
  { key: 'Unpaid', text: 'Unpaid' },
  { key: 'Work From Home', text: 'Work From Home' }
];

function isSameDay(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

export function ApplyLeaveForm(props) {
  const [leaveType, setLeaveType] = React.useState('Casual');
  const [startDate, setStartDate] = React.useState(undefined);
  const [endDate, setEndDate] = React.useState(undefined);
  const [reason, setReason] = React.useState('');
  const [isHalfDay, setIsHalfDay] = React.useState(false);
  const [attachment, setAttachment] = React.useState(undefined);

  const [error, setError] = React.useState('');
  const [success, setSuccess] = React.useState('');
  const [submitting, setSubmitting] = React.useState(false);

  const canHalfDay = !!startDate && !!endDate && isSameDay(startDate, endDate);

  React.useEffect(() => {
    if (!canHalfDay && isHalfDay) {
      setIsHalfDay(false);
    }
  }, [canHalfDay, isHalfDay]);

  const validate = () => {
    if (!startDate) return 'Start date is required.';
    if (!endDate) return 'End date is required.';
    if (endDate.getTime() < startDate.getTime()) return 'End date must be same or after start date.';

    if (leaveType === 'Sick' && isSameDay(startDate, endDate)) {
      if (!attachment) return 'Attachment is required for 1-day Sick Leave.';
    }

    if (isHalfDay && !isSameDay(startDate, endDate)) {
      return 'Half day leave is allowed only for a single day.';
    }

    return '';
  };

  const onSubmit = async () => {
    setError('');
    setSuccess('');

    const validationError = validate();
    if (validationError) {
      setError(validationError);
      return;
    }

    setSubmitting(true);
    try {
      const id = await props.service.applyLeave({
        leaveType,
        startDate,
        endDate,
        reason,
        isHalfDay,
        attachment
      });

      setSuccess(`Leave request submitted. ID: ${id}`);
      setReason('');
      setIsHalfDay(false);
      setAttachment(undefined);
      props.onSubmitted();
    } catch (e) {
      setError(e && e.message ? e.message : 'Failed to submit leave request.');
    } finally {
      setSubmitting(false);
    }
  };

  const onReset = () => {
    setError('');
    setSuccess('');
    setLeaveType('Casual');
    setStartDate(undefined);
    setEndDate(undefined);
    setReason('');
    setIsHalfDay(false);
    setAttachment(undefined);
  };

  return React.createElement(
    Stack,
    { tokens: { childrenGap: 12 } },
    error ? React.createElement(MessageBar, { messageBarType: MessageBarType.error }, error) : null,
    success ? React.createElement(MessageBar, { messageBarType: MessageBarType.success }, success) : null,
    React.createElement(Dropdown, {
      label: 'Leave Type',
      options: leaveTypeOptions,
      selectedKey: leaveType,
      onChange: (_, opt) => setLeaveType(String(opt && opt.key))
    }),
    React.createElement(DatePicker, {
      label: 'Start Date',
      firstDayOfWeek: DayOfWeek.Monday,
      value: startDate,
      onSelectDate: (d) => setStartDate(d || undefined),
      placeholder: 'Select a date'
    }),
    React.createElement(DatePicker, {
      label: 'End Date',
      firstDayOfWeek: DayOfWeek.Monday,
      value: endDate,
      onSelectDate: (d) => setEndDate(d || undefined),
      placeholder: 'Select a date'
    }),
    React.createElement(Toggle, {
      label: 'Half Day',
      checked: isHalfDay,
      onChange: (_, checked) => setIsHalfDay(!!checked),
      disabled: !canHalfDay,
      onText: 'Yes',
      offText: 'No'
    }),
    React.createElement(TextField, {
      label: 'Reason',
      multiline: true,
      rows: 3,
      value: reason,
      onChange: (_, v) => setReason(v || '')
    }),
    React.createElement(
      'div',
      null,
      React.createElement('label', null, 'Attachment'),
      React.createElement('input', {
        type: 'file',
        onChange: (e) => setAttachment(e.target.files && e.target.files.length > 0 ? e.target.files[0] : undefined)
      })
    ),
    React.createElement(
      Stack,
      { horizontal: true, tokens: { childrenGap: 10 } },
      React.createElement(PrimaryButton, { text: 'Submit', onClick: onSubmit, disabled: submitting }),
      React.createElement(DefaultButton, { text: 'Reset', onClick: onReset, disabled: submitting }),
      submitting ? React.createElement(Spinner, { size: SpinnerSize.small }) : null
    )
  );
}
