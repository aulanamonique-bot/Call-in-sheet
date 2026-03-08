exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ok: false, error: 'Method not allowed' })
    };
  }

  try {
    const body = JSON.parse(event.body || '{}');

    const appId = process.env.APPSHEET_APP_ID;
    const tableName = process.env.APPSHEET_TABLE_NAME;
    const apiKey = process.env.APPSHEET_API_KEY;

    if (!appId || !tableName || !apiKey) {
      return {
        statusCode: 500,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ok: false, error: 'Missing AppSheet environment variables' })
      };
    }

    const timestampValue = buildTimestamp(body.timestampDate, body.timestampTime);

    const row = {
      "Timestamp": timestampValue,
      "Notes": body.notes || "",
      "Source": body.source || "",
      "Customer Name": body.customerName || "",
      "Phone": body.phone || "",
      "Email": body.email || "",
      "Address": body.address || "",
      "Interior": body.interior || "N",
      "Power Spray": body.powerSpray || "N",
      "Web Dust": body.webDust || "N",
      "Granulate": body.granulate || "N",
      "Leaf Blow": body.leafBlow || "N",
      "Reservices": body.reservices || "N",
      "Services Needed": body.servicesNeeded || "",
      "Quote Initial": body.quoteInitial === "" ? "" : Number(body.quoteInitial || 0),
      "Quote QT": body.quoteQT === "" ? "" : Number(body.quoteQT || 0),
      "Liquid Upgrade": body.liquidUpgrade || "N",
      "Termite Install Fee": body.termiteInstallFee === "" ? "" : Number(body.termiteInstallFee || 0),
      "Termite Annual": body.termiteAnnual === "" ? "" : Number(body.termiteAnnual || 0),
      "Payment Type": body.paymentType || "",
      "Status": body.status || "Pending",
      "Audited": body.audited || "N"
    };

    const appSheetRes = await fetch(
      `https://api.appsheet.com/api/v2/apps/${encodeURIComponent(appId)}/tables/${encodeURIComponent(tableName)}/Action`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'ApplicationAccessKey': apiKey
        },
        body: JSON.stringify({
          Action: 'Add',
          Properties: {
            Locale: 'en-US',
            Timezone: 'America/New_York'
          },
          Rows: [row]
        })
      }
    );

    const text = await appSheetRes.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      data = { raw: text };
    }

    if (!appSheetRes.ok) {
      return {
        statusCode: appSheetRes.status,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ok: false,
          error: data?.error || data?.Message || 'AppSheet request failed',
          details: data
        })
      };
    }

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ok: true, data })
    };
  } catch (error) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ok: false, error: error.message || 'Unknown error' })
    };
  }
};

function buildTimestamp(dateStr, timeStr) {
  if (!dateStr) return new Date().toISOString();

  const safeTime = timeStr || '00:00';
  const isoLike = `${dateStr}T${safeTime}:00`;
  const d = new Date(isoLike);

  if (Number.isNaN(d.getTime())) {
    return new Date().toISOString();
  }

  return d.toISOString();
}
