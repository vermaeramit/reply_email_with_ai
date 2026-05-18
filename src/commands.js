const WEBHOOK_URL = 'https://ai.worxpertise.com/webhook/c9858c5b-352c-4a90-8dca-5e800add4e7e';

Office.onReady(() => {
  Office.actions.associate('generateReply', generateReply);
  Office.actions.associate('generateReplyFromRead', generateReplyFromRead);
});

async function generateReply(event) {
  await runReplyFlow(event, /* fromRead */ false);
}

async function generateReplyFromRead(event) {
  await runReplyFlow(event, /* fromRead */ true);
}

async function runReplyFlow(event, fromRead) {
  try {
    const item = Office.context.mailbox.item;
    if (!item) return finish(event, 'No active mail item.');

    notify('aier-status', 'informationalMessage', 'Generating reply...');

    if (fromRead) {
      await new Promise((resolve, reject) => {
        item.displayReplyAllForm({ htmlBody: '' });
        setTimeout(resolve, 1200);
      });
      return finish(event, null, 'Reply opened. Click the AI Reply button in the compose ribbon to generate text.');
    }

    const subject = await getSubject(item);
    const thread = await getBody(item);

    const res = await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ subject, thread })
    });
    const raw = await res.text();
    if (!res.ok) return finish(event, 'Webhook ' + res.status + ': ' + raw.slice(0, 160));

    let data = null;
    try { data = JSON.parse(raw); } catch {}
    const reply = data
      ? (data.reply ?? data.output ?? data.text ?? data.message ?? data.body ?? raw)
      : raw;

    await insertReply(item, String(reply));
    finish(event, null, 'Reply inserted.');
  } catch (err) {
    finish(event, 'Error: ' + (err && err.message ? err.message : String(err)));
  }
}

function getSubject(item) {
  return new Promise((resolve) => {
    try {
      if (typeof item.subject === 'string') return resolve(item.subject);
      if (item.subject && typeof item.subject.getAsync === 'function') {
        return item.subject.getAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || '') : ''));
      }
      resolve('');
    } catch {
      resolve('');
    }
  });
}

function getBody(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (r) => {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value || '') : '');
    });
  });
}

function insertReply(item, text) {
  return new Promise((resolve, reject) => {
    const payload = (text || '').replace(/\s+$/, '') + '\n\n';
    item.body.setSelectedDataAsync(payload, { coercionType: Office.CoercionType.Text }, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) return resolve();
      // Fallback: prepend at cursor via prependAsync if available
      if (item.body.prependAsync) {
        item.body.prependAsync(payload, { coercionType: Office.CoercionType.Text }, (r2) => {
          if (r2.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(r2.error && r2.error.message ? r2.error.message : 'Insert failed'));
        });
      } else {
        reject(new Error(r.error && r.error.message ? r.error.message : 'Insert failed'));
      }
    });
  });
}

function notify(key, type, message) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync(key, {
      type,
      message: String(message).slice(0, 150),
      icon: type === 'informationalMessage' ? 'Icon.16x16' : undefined,
      persistent: false
    });
  } catch {}
}

function finish(event, errorMessage, infoMessage) {
  if (errorMessage) notify('aier-status', 'errorMessage', errorMessage);
  else if (infoMessage) notify('aier-status', 'informationalMessage', infoMessage);
  event.completed();
}
