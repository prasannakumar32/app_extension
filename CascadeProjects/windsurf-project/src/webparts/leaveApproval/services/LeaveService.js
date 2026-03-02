import { SPHttpClient } from '@microsoft/sp-http';

export class LeaveService {
  constructor(webAbsoluteUrl, spHttpClient, listTitle) {
    this._webAbsoluteUrl = webAbsoluteUrl;
    this._spHttpClient = spHttpClient;
    this._listTitle = listTitle;
  }

  async _getCurrentUser() {
    const url = `${this._webAbsoluteUrl}/_api/web/currentuser?$select=Id,Email,Title`;
    const res = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata.metadata=none' }
    });

    if (!res.ok) {
      throw new Error(`Failed to load current user. ${res.status} ${res.statusText}`);
    }

    return await res.json();
  }

  _toDateOnlyIso(d) {
    const pad2 = (n) => (n < 10 ? `0${n}` : String(n));
    const yyyy = d.getFullYear();
    const mm = pad2(d.getMonth() + 1);
    const dd = pad2(d.getDate());
    return `${yyyy}-${mm}-${dd}T00:00:00Z`;
  }

  async applyLeave(request) {
    const user = await this._getCurrentUser();

    const createUrl = `${this._webAbsoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this._listTitle)}')/items`;

    const body = {
      Title: `Leave - ${new Date().toISOString()}`,
      LeaveType: request.leaveType,
      StartDate: this._toDateOnlyIso(request.startDate),
      EndDate: this._toDateOnlyIso(request.endDate),
      Reason: request.reason,
      IsHalfDay: request.isHalfDay,
      Status: 'Pending',
      EmployeeNameId: user.Id
    };

    const res = await this._spHttpClient.post(createUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata.metadata=none',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Failed to create leave request. ${res.status} ${res.statusText} ${text}`);
    }

    const created = await res.json();

    if (request.attachment) {
      await this.addAttachment(created.Id, request.attachment);
    }

    return created.Id;
  }

  async addAttachment(itemId, file) {
    const fileBuffer = await file.arrayBuffer();
    const uploadUrl = `${this._webAbsoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this._listTitle)}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;

    const res = await this._spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata.metadata=none',
        'Content-Type': 'application/octet-stream'
      },
      body: fileBuffer
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Failed to upload attachment. ${res.status} ${res.statusText} ${text}`);
    }
  }

  async getMyLeaves() {
    const user = await this._getCurrentUser();

    const baseUrl = `${this._webAbsoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this._listTitle)}')/items`;
    const select = '$select=Id,LeaveType,StartDate,EndDate,Reason,IsHalfDay,Status,ApproverComments,Created';
    const order = '$orderby=Created desc';

    const filter = `EmployeeNameId eq ${user.Id}`;
    const url = `${baseUrl}?${select}&$filter=${filter}&${order}`;

    const res = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata.metadata=none' }
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Failed to load leave history. ${res.status} ${res.statusText} ${text}`);
    }

    const json = await res.json();
    return json.value;
  }
}
