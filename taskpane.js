(() => {
"use strict";
const ENDPOINT = "http://127.0.0.1:8765/review";
const button = document.getElementById("review");
const status = document.getElementById("status");
function update(message) { status.textContent = message; }
Office.onReady(info => {
 if (info.host !== Office.HostType.Word) { update("Open in Microsoft Word."); return; }
 if (!Office.context.requirements.isSetSupported("WordApi","1.4")) { update("WordApi 1.4 is required for comments."); return; }
 button.disabled = false; update("Ready to review selected text or the document.");
});
button.addEventListener("click", async () => {
 button.disabled = true;
 try {
   update("Reading report…");
   const report = await Word.run(async ctx => {
     const selection = ctx.document.getSelection(), body = ctx.document.body;
     selection.load("text"); body.load("text"); await ctx.sync();
     const s = selection.text.trim(), text = s || body.text.trim();
     if (!text || text.length > 18000) throw Error("Report must contain 1–18,000 characters.");
     return {text, selected: !!s};
   });
   update("Reviewing with Gemini…");
   const controller = new AbortController(), timeout = setTimeout(() => controller.abort(), 90000);
   let data;
   try {
     const res = await fetch(ENDPOINT, {method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({text:report.text}),signal:controller.signal,cache:"no-store"});
     data = await res.json().catch(() => ({}));
     if (!res.ok) throw Error(data.error || "Local service failed (HTTP " + res.status + ").");
   } finally {clearTimeout(timeout);}
   const issues = Array.isArray(data.issues) ? data.issues.slice(0,30) : [];
   let annotated=0, skipped=0;
   update("Annotating " + issues.length + " suggestions…");
   await Word.run(async ctx => {
     const results=[];
     for(const issue of issues) {
       const quote = String(issue.quote||"").trim();
       if(!quote || quote.length > 300 || !report.text.includes(quote)) {skipped++;continue;}
       const search = report.selected ? ctx.document.getSelection().search(quote,{matchCase:true}) : ctx.document.body.search(quote,{matchCase:true});
       search.load("items");results.push({search,issue});
     }
     await ctx.sync();
     for(const {search,issue} of results) {
       if(search.items.length!==1){skipped++;continue;}
       const range=search.items[0];
       range.font.highlightColor=issue.category==="logic"?"#FF5555":"#FFF176";
       const explanation=String(issue.explanation||"Verify this passage.").slice(0,650);
       const suggestion=String(issue.suggestion||"").slice(0,350);
       range.insertComment("[AI suggestion – " + (issue.category==="logic"?"Potential internal inconsistency":"Language") + "]\n" + explanation + (suggestion?"\nSuggested: "+suggestion:"") + "\nClinical verification required.");
       annotated++;
     }
     await ctx.sync();
   });
   update("Review complete: "+annotated+" annotated; "+skipped+" skipped. Yellow = language, red = potential internal contradiction. Nothing was rewritten.");
 } catch(err) { update("Error: "+(err.name==="AbortError"?"Request timed out.":err.message||String(err))); }
 finally {button.disabled=false;}
});
})();