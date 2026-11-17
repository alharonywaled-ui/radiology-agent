using System;
using System.Text.RegularExpressions;

namespace RadiologyAgent {
  // Pure text planner: no COM side effects. The scope shown in the popup
  // MUST be the scope replaced. Never paste a complete sentence into a
  // two-word anchor.
  internal static class SuggestionReplacementPlan {
    internal static bool TryBuild(string report,int anchorStart,int anchorEnd,
       string quote,string proposal,out int start,out int end,out string original,
       out string safeProposal) {
      start=anchorStart;end=anchorEnd;original=quote;safeProposal="";
      if(String.IsNullOrEmpty(report)||String.IsNullOrWhiteSpace(quote)||
         anchorStart<0||anchorEnd>report.Length||anchorEnd<=anchorStart||
         !String.Equals(report.Substring(anchorStart,anchorEnd-anchorStart),quote,
                        StringComparison.Ordinal))return false;
      string candidate=(proposal??"").Trim();
      if(!SuggestionPopup.IsLiteralReplacement(candidate))return true;
      if(candidate.Length>2000||candidate.IndexOf('\r')>=0||
         candidate.IndexOf('\n')>=0||candidate.IndexOf('\x07')>=0)return true;
      // A direct, local spelling replacement. Longer suggestions are NOT
      // assumed to replace a short quoted fragment.
      int originalWords=Words(quote),newWords=Words(candidate);
      if(newWords<=originalWords+2 &&
         candidate.Length<=Math.Max(quote.Length+26,quote.Length*2)) {
        safeProposal=candidate;return true;
      }
      // For complete-sentence suggestions, first locate the actual sentence
      // around the error, not the abbreviated error quote.
      int left=anchorStart,right=anchorEnd;
      while(left>0&&!IsBoundary(report,left-1))left--;
      while(right<report.Length&&!IsBoundary(report,right))right++;
      if(right<report.Length&&(report[right]=='.'||report[right]=='?'||report[right]=='!'))
        right++;
      // Never cross Word paragraphs, tables, signatures or other text blocks.
      while(left<anchorStart&&Char.IsWhiteSpace(report[left]))left++;
      while(right>anchorEnd&&Char.IsWhiteSpace(report[right-1]))right--;
      if(left>=right||left>anchorStart||right<anchorEnd)return true;
      string full=report.Substring(left,right-left);
      if(full.Length>500||full.Length<quote.Length+12)return true;
      string normalizedOld=Normalize(full),normalizedNew=Normalize(candidate);
      if(normalizedOld==normalizedNew)return true;
      // Require substantial intact context to avoid unintended clinical
      // rewrites. This detects one-word corrections in full sentences but
      // rejects unrelated sentences that happen to share a single word.
      int prefix=CommonPrefix(normalizedOld,normalizedNew);
      int suffix=CommonSuffix(normalizedOld,normalizedNew,prefix);
      int shorter=Math.Min(normalizedOld.Length,normalizedNew.Length);
      bool enough=(prefix+suffix>=Math.Max(20,(int)(shorter*0.52))) &&
                  (prefix>=10||suffix>=14) &&
                  (normalizedNew.Length>=Math.Max(20,normalizedOld.Length/2)) &&
                  (normalizedNew.Length<=normalizedOld.Length*2);
      if(enough) {
        start=left;end=right;original=full;safeProposal=candidate;
      }
      return true;
    }

    private static bool IsBoundary(string s,int i) {
      char c=s[i];
      if(c=='\r'||c=='\n'||c=='\x07'||c=='?'||c=='!')return true;
      if(c=='.') {
        // Decimal points (e.g. 3.5 cm) must not split the sentence.
        if(i>0&&i+1<s.Length&&Char.IsDigit(s[i-1])&&Char.IsDigit(s[i+1]))return false;
        return i+1==s.Length||Char.IsWhiteSpace(s[i+1]);
      }
      return false;
    }
    private static int Words(string s) {
      return Regex.Matches(s??"", @"\b[\p{L}\p{N}]+\b").Count;
    }
    private static string Normalize(string s) {
      return Regex.Replace((s??"").Trim(),@"\s+"," ");
    }
    private static int CommonPrefix(string a,string b) {
      int n=Math.Min(a.Length,b.Length),i=0;
      while(i<n&&Char.ToLowerInvariant(a[i])==Char.ToLowerInvariant(b[i]))i++;
      return i;
    }
    private static int CommonSuffix(string a,string b,int prefix) {
      int n=Math.Min(a.Length,b.Length)-prefix,i=0;
      while(i<n&&Char.ToLowerInvariant(a[a.Length-1-i])==
                   Char.ToLowerInvariant(b[b.Length-1-i]))i++;
      return i;
    }
    internal static bool SelfTest() {
      string sentence="Elevated left diaphragmatic copula with underlying lower lobar thick atelectatic parenchymal bands.";
      string correction="Elevated left diaphragmatic dome with underlying lower lobar thick atelectatic parenchymal bands.";
      string doc="Findings:\r"+sentence+"\rConclusion:\rNo stone.";
      int at=doc.IndexOf("diaphragmatic copula",StringComparison.Ordinal);
      int start,end;string original,proposal;
      if(!TryBuild(doc,at,at+"diaphragmatic copula".Length,
                  "diaphragmatic copula",correction,
                  out start,out end,out original,out proposal)||
         original!=sentence||proposal!=correction||
         start!=doc.IndexOf(sentence,StringComparison.Ordinal))return false;
      if(!TryBuild(doc,at,at+"diaphragmatic copula".Length,
                  "diaphragmatic copula","diaphragmatic dome",
                  out start,out end,out original,out proposal)||
         original!="diaphragmatic copula"||proposal!="diaphragmatic dome")return false;
      if(!TryBuild(doc,at,at+"diaphragmatic copula".Length,
                  "diaphragmatic copula","Entirely unrelated replacement sentence about another organ.",
                  out start,out end,out original,out proposal)||
         original!="diaphragmatic copula"||proposal!="")return false;
      if(TryBuild(doc,at+1,at+3,"diaphragmatic copula",correction,
                  out start,out end,out original,out proposal))return false;
      return true;
    }
  }
}
