using System;
using System.Text.RegularExpressions;

namespace RadiologyAgent {
  // Pure, Word-independent replacement planning. A full-sentence suggestion
  // must replace the entire original sentence, never a two-word comment scope.
  internal static class SuggestionPlan {
    internal sealed class Plan {
      internal readonly int Start,End;
      internal readonly string Original,Proposed;
      internal Plan(int start,string original,string proposed) {
        Start=start;End=start+original.Length;Original=original;Proposed=proposed;
      }
    }
    internal static bool IsAdvice(string suggestion) {
      if(String.IsNullOrWhiteSpace(suggestion))return true;
      string low=suggestion.Trim().ToLowerInvariant();
      string[] prefixes={"suggestion:","verify ","confirm ","review ","ask ","consider ",
        "check ","clarify ","correlate ","ensure ","possibly ","perhaps ",
        "if ","may ","could ","should ","requires ","either ","add supporting",
        "reconcile ","depending ","it is ","the reporting radiologist"};
      foreach(string p in prefixes)if(low.StartsWith(p,StringComparison.Ordinal))return true;
      return suggestion.Contains("\r")||suggestion.Contains("\n");
    }

    internal static Plan Build(string quote,string proposal,string paragraph,
                               int quoteOffset,int paragraphAbsoluteStart) {
      if(String.IsNullOrWhiteSpace(quote)||String.IsNullOrWhiteSpace(proposal)||
         String.IsNullOrEmpty(paragraph)||quoteOffset<0||
         quoteOffset+quote.Length>paragraph.Length||
         !String.Equals(paragraph.Substring(quoteOffset,quote.Length),quote,
                        StringComparison.Ordinal)||IsAdvice(proposal))
        return null;
      string proposed=proposal.Trim();
      if(proposed.Length>2000)return null;
      // Trim Word table-cell markers and paragraph returns from either edge.
      int localStart=quoteOffset,localEnd=quoteOffset+quote.Length;
      int sentenceStart=quoteOffset,sentenceEnd=localEnd;
      while(sentenceStart>0) {
        char c=paragraph[sentenceStart-1];
        if(c=='\r'||c=='\n'||c=='\x07'||c==';'||c=='!'||c=='?'||
           (c=='.'&&(sentenceStart>=paragraph.Length||paragraph[sentenceStart]!=' ')))
          break;
        if(c=='.'&&sentenceStart<paragraph.Length&&paragraph[sentenceStart]==' ')break;
        sentenceStart--;
      }
      while(sentenceStart<quoteOffset&&Char.IsWhiteSpace(paragraph[sentenceStart]))sentenceStart++;
      while(sentenceEnd<paragraph.Length) {
        char c=paragraph[sentenceEnd];
        if(c=='\r'||c=='\n'||c=='\x07')break;
        sentenceEnd++;
        if(c=='.'||c=='!'||c=='?')break;
      }
      while(sentenceEnd>sentenceStart&&Char.IsWhiteSpace(paragraph[sentenceEnd-1]))sentenceEnd--;
      string sentence=paragraph.Substring(sentenceStart,sentenceEnd-sentenceStart);
      if(sentence.Length<quote.Length)return null;
      int relativeQuote=quoteOffset-sentenceStart;
      string prefix=sentence.Substring(0,relativeQuote);
      string suffix=sentence.Substring(relativeQuote+quote.Length);
      string prefixTrim=prefix.Trim(),suffixTrim=suffix.Trim();

      // If replacement is a sentence and retains the surrounding context,
      // treat it as replacing the entire sentence. Never stuff it into quote.
      bool hasPrefix=prefixTrim.Length>=5&&
          proposed.StartsWith(prefixTrim,StringComparison.OrdinalIgnoreCase);
      bool hasSuffix=suffixTrim.Length>=5&&
          proposed.EndsWith(suffixTrim,StringComparison.OrdinalIgnoreCase);
      bool longProposal=proposed.Length>quote.Length+Math.Max(18,quote.Length/2);
      bool wholeSentence=longProposal&&(hasPrefix||hasSuffix);
      if(wholeSentence) {
        if(String.Equals(sentence,proposed,StringComparison.Ordinal))return null;
        return new Plan(paragraphAbsoluteStart+sentenceStart,sentence,proposed);
      }
      // A long, mismatching proposed sentence may not safely replace a short
      // quote. No unsafe auto-Apply; offer manual editing instead.
      if(longProposal)return null;
      if(String.Equals(quote,proposed,StringComparison.Ordinal))return null;
      return new Plan(paragraphAbsoluteStart+localStart,quote,proposed);
    }
    internal static bool SelfTest() {
      const string original="Elevated left diaphragmatic copula with underlying lower lobar thick atelectatic parenchymal bands.";
      const string proposed="Elevated left diaphragmatic dome with underlying lower lobar thick atelectatic parenchymal bands.";
      int q=original.IndexOf("diaphragmatic copula",StringComparison.Ordinal);
      Plan full=Build("diaphragmatic copula",proposed,original,q,100);
      Plan phrase=Build("diaphragmatic copula","diaphragmatic dome",original,q,100);
      Plan bad=Build("diaphragmatic copula","Different unrelated report statement with no surrounding context.",original,q,100);
      Plan advisory=Build("diaphragmatic copula","Verify the anatomical findings.",original,q,100);
      return full!=null&&full.Start==100&&full.Original==original&&full.Proposed==proposed&&
        phrase!=null&&phrase.Start==100+q&&phrase.Original=="diaphragmatic copula"&&
        phrase.Proposed=="diaphragmatic dome"&&bad==null&&advisory==null;
    }
  }
}
