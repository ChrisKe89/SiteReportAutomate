website:
https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx

elements
<div id="pnlSearch">
                            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tbody><tr>
                                    <td width="10%">
                                        <span id="lblOpCO">OpCo :</span><span class="star"> *</span>
                                    </td>
                                    <td width="15%">
                                        <select name="ctl00$MainContent$ddlOpCoID" onchange="javascript:setTimeout('__doPostBack(\'ctl00$MainContent$ddlOpCoID\',\'\')', 0)" id="MainContent_ddlOpCoID" class="yellowBg">
	<option value="Select">Select</option>
	<option selected="selected" value="FXAU">FBAU</option>
	<option value="FXCA">FBCA</option>
	<option value="FXID">FBCN</option>
	<option value="FXHK">FBHK</option>
	<option value="FXKR">FBKR</option>
	<option value="FXMM">FBMM</option>
	<option value="FXMY">FBMY</option>
	<option value="FXNZ">FBNZ</option>
	<option value="FXPH">FBPH</option>
	<option value="FXSG">FBSG</option>
	<option value="THFX">FBTH</option>
	<option value="FXTW">FBTW</option>
	<option value="FXVN">FBVN</option>
	<option value="TWSI">TWSI</option>

</select>
                                        
                                    </td>

                                    <td width="10%">Product Code :<span class="star">  *</span>
                                    </td>
                                    <td width="15%">
                                        <input name="ctl00$MainContent$ProductCode" type="text" value="TC101632" id="MainContent_ProductCode">
                                    </td>

                                    <td width="10%">Serial Number :<span class="star"> *</span>
                                    </td>
                                    <td width="15%">
                                        <input name="ctl00$MainContent$SerialNumber" type="text" value="131586" id="MainContent_SerialNumber">
                                    </td>

                                </tr>
                                <tr>
                                    <td colspan="5" class="mandatory">* Mandatory
                                    </td>
                                    <td align="right" class="buttonRight">
                                        <input type="submit" name="ctl00$MainContent$btnSearch" value="Search" id="MainContent_btnSearch" class="btn btn-small btn-info button-small">
                                        <input type="submit" name="ctl00$MainContent$btnReset" value="Reset" id="MainContent_btnReset" class="btn btn-small btn-info button-small">
                                    </td>
                                </tr>
                            </tbody></table>
                        </div>

change this to FBAU:
<select name="ctl00$MainContent$ddlOpCoID" onchange="javascript:setTimeout('__doPostBack(\'ctl00$MainContent$ddlOpCoID\',\'\')', 0)" id="MainContent_ddlOpCoID" class="yellowBg">
	<option value="Select">Select</option>
	<option selected="selected" value="FXAU">FBAU</option>
	<option value="FXCA">FBCA</option>
	<option value="FXID">FBCN</option>
	<option value="FXHK">FBHK</option>
	<option value="FXKR">FBKR</option>
	<option value="FXMM">FBMM</option>
	<option value="FXMY">FBMY</option>
	<option value="FXNZ">FBNZ</option>
	<option value="FXPH">FBPH</option>
	<option value="FXSG">FBSG</option>
	<option value="THFX">FBTH</option>
	<option value="FXTW">FBTW</option>
	<option value="FXVN">FBVN</option>
	<option value="TWSI">TWSI</option>

</select>


enter product code from xlsx here:
<input name="ctl00$MainContent$ProductCode" type="text" value="" id="MainContent_ProductCode">

enter serial number from xlsx here:
<input name="ctl00$MainContent$SerialNumber" type="text" value="131586" id="MainContent_SerialNumber">

search:
<input type="submit" name="ctl00$MainContent$btnSearch" value="Search" id="MainContent_btnSearch" class="btn btn-small btn-info button-small">


if this appears
<table class="listing removeBottonMargin tblAlertApply" cellspacing="0" rules="all" border="1" id="MainContent_GridViewEligibility" style="width:100%;border-collapse:collapse;">
			<tbody><tr>
				<th scope="col">OpCo</th><th scope="col">Product Code</th><th scope="col">Serial Number</th><th scope="col">Product Family</th><th scope="col">Registration Status</th><th scope="col">Service Type</th><th scope="col">Software Upgrade Capability</th><th scope="col">Reserved Software Upgrade Capability</th><th scope="col">Reserved Software Upgrade Flag</th><th scope="col">Software Upgrade By KO Flag</th><th scope="col">Status for Software Upgrade By KO</th>
			</tr><tr>
				<td>FXAU</td><td>TC101632</td><td>131586</td><td>Greif</td><td>epregistered</td><td>58</td><td style="background-color:Yellow;">Enabled</td><td style="background-color:Yellow;">Enabled</td><td style="background-color:Yellow;">Enabled</td><td style="background-color:Yellow;">Enabled</td><td style="background-color:Yellow;">already upgraded</td>
			</tr>
		</tbody></table>

and this says already upgraded move to next row:
<td style="background-color:Yellow;">already upgraded</td>

if this appears:

<table class="listing removeBottonMargin tblAlertApply" cellspacing="0" rules="all" border="1" id="MainContent_GridViewDevice" style="width:100%;border-collapse:collapse;">
			<tbody><tr>
				<th scope="col">OpCo</th><th scope="col">Product Code</th><th scope="col">Serial Number</th><th scope="col">Product Family</th><th scope="col">Service Type</th><th scope="col">Registration Status</th><th scope="col">Software Upgrade Capability</th>
			</tr><tr>
				<td>FBAU</td><td>TC101632</td><td>131040</td><td>Greif</td><td>58</td><td>epregistered</td><td>true</td>
			</tr>
		</tbody></table>


enter a random date in here using date picker (can only be 6 days from today)
<input name="ctl00$MainContent$txtDateTime" type="text" maxlength="50" id="MainContent_txtDateTime" class="txtDateTime" readonly="readonly" style="width: 75%;">

select random time 12am to 7am or 6pm to 11pm

choose this based on state column NT = "Darwin", SA = "Adelaide", ACT/VIC/NSW = "Canberra, Melbourne, Sydney", QLD = "Brisbane", TAS = "Hobart"
<select name="ctl00$MainContent$ddlTimeZone" id="MainContent_ddlTimeZone" class="yellowBg">
		<option selected="selected" value="Select">Select Time Zone</option>
		<option value="+09:30">(UTC+09:30) Darwin</option>
		<option value="+10:30">(UTC+10:30) Adelaide</option>
		<option value="+11:00">(UTC+11:00) Canberra, Melbourne, Sydney</option>
		<option value="+10:00">(UTC+10:00) Brisbane</option>
		<option value="+11:00">(UTC+11:00) Hobart</option>

	</select>

Click Schedule
<input type="submit" name="ctl00$MainContent$submitButton" value="Schedule" id="MainContent_submitButton" class="btn btn-small btn-info button-medium">