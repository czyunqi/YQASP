<!--#include file="_cls.testcase.asp" --><%
Dim rs
Set rs = YQasp.Db.Query("Select ContentID As id, ContentClassID As cid, ContentTitle As title, AnnounceTime As atime, ContentText As content From EC_Content Where ContentID = {id}")
YQasp.Print YQasp.Encode(rs)
YQasp.Db.Close(rs)
%>