export const CONTACT_SHEET_URL='https://docs.google.com/spreadsheets/d/1cRcI4qp5daSaaZ1OvEjZti13IZHX0Uz-nr1M35t_58s/edit?gid=0#gid=0';
export const CONTACT_EXPORT_URL='https://docs.google.com/spreadsheets/d/1cRcI4qp5daSaaZ1OvEjZti13IZHX0Uz-nr1M35t_58s/export?format=csv&gid=0';
export async function fetchContactSheet(fetcher=fetch){
 const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),25000);
 try{
  // No staff token or cookies are sent to Google. This is a read-only download.
  const response=await fetcher(CONTACT_EXPORT_URL,{credentials:'omit',cache:'no-store',referrerPolicy:'no-referrer',signal:controller.signal});
  if(!response.ok)throw Error('The contact sheet is not accessible with its current permissions. Keep it restricted if required and import an Excel download instead.');
  const type=response.headers.get('content-type')||'';if(!/text\/(csv|plain)|application\/csv/i.test(type))throw Error('Google returned a sign-in page instead of patient data. Check sheet access or import an Excel download.');
  const csv=await response.text();if(new TextEncoder().encode(csv).length>30*1024*1024)throw Error('The sheet is too large for this update. Import it in smaller files.');
  if(!csv.trim()||/^\s*<!?html/i.test(csv))throw Error('The contact sheet did not return usable CSV data.');
  return {csv,sourceUrl:CONTACT_SHEET_URL,fetchedAt:new Date().toISOString()};
 }catch(error){if(error.name==='AbortError')throw Error('The sheet download timed out. No patient data was changed. Try again after reconnection.');throw error;}finally{clearTimeout(timer);}
}
