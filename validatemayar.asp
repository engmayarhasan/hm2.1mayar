<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Validate Information</title>
    <style>
        main{
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            min-height: 500px;
            background-color: whitesmoke;
        }
        .data{
            display: flex;
            flex-direction: column;
        }
    </style>
</head>
<body>
    <main>
        <h2>You gave us your data!</h2>
        <div class="data">
            <%
                dim fullName
                fullName=Request.Form("fullName")
                dim card
                card=Request.Form("card")
                dim section
                section=Request.Form("section")
                dim company
                company=Request.Form("company")
            %>

            <p>
                <%
                Response.Write("Your Full Name is :" & fullName)
                %>
            </p>
            <p>
                <%
                Response.Write("Your Card Number is :" & card)
                %>
            </p>
            <p>
                <%
                Response.Write("Your Section is :" & section)
                %>
            </p>
            <p>
                <%
                Response.Write("Your Card Company is :" & company)
                %>
            </p>
        </div>
        <div class="file">
            <h2>Text File Content</h2>
            <%
                Set fs=Server.CreateObject("Scripting.FileSystemObject")

                If (fs.FileExists("C:\inetpub\wwwroot\aspForms\validate.txt"))=true Then
                    Set f=fs.OpenTextFile(Server.MapPath("validate.txt"), 1)
                    Response.Write(f.ReadAll)
                    f.Close
                    
                End If
                Set f=Nothing
                set fs=nothing
            %>
        </div>
    </main>
</body>
</html>