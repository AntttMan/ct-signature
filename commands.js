/* Combustion Technology — email signature add-in (event handler) */

var BASE_URL = "https://antttman.github.io/ct-signature";

Office.actions.associate("onNewMessageCompose", onNewMessageCompose);

function onNewMessageCompose(event) {
  var raw = Office.context.roamingSettings.get("ct_signature_info");
  if (!raw) {
    Office.context.mailbox.item.notificationMessages.addAsync("ct-sig-setup", {
      type: "informationalMessage",
      message: "Click the CT Signature button in the ribbon to set up your email signature.",
      icon: "Icon.16x16",
      persistent: false
    }, function () { event.completed(); });
    return;
  }

  var info = JSON.parse(raw);
  var html = buildSignature(info);

  Office.context.mailbox.item.disableClientSignatureAsync(function () {
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: "html" },
      function () { event.completed(); }
    );
  });
}

function buildSignature(info) {
  var name = esc(info.name);
  var designation = esc(info.designation);
  var cell = esc(info.cell);
  var ext = esc(info.extension);
  var email = esc(info.email || Office.context.mailbox.userProfile.emailAddress);
  var mailtoEmail = info.email || Office.context.mailbox.userProfile.emailAddress;

  return [
    '<div style="font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;color:#000;">Kind regards,</div>',
    '<div style="font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;color:#000;"><br></div>',
    '<div style="background-color:#fff;margin:0;">',

    // Name & title
    '<p style="margin:0;"><span style="font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;color:#000;"><b>' + name + '</b></span></p>',
    '<div style="font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;color:#000;"><b>' + designation + '</b></div>',
    '<div style="font-size:10pt;"><b><br></b></div>',

    // Logo
    '<div style="font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;color:#000;">',
    '<img alt="Combustion Technology" width="253" height="73" src="' + BASE_URL + '/assets/logo.png" style="width:253px;height:73px;"></div>',

    // Tagline
    '<div style="font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;color:#000;"><b>LEAD | TRUST | RESPOND | PROTECT</b></div>',
    '<div style="font-size:10pt;"><br></div>',

    // Contact details
    ext ? '<p style="line-height:normal;margin:0;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;">' +
    '<span style="color:#000;"><b>Tel: &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </b></span>' +
    '<span style="color:rgb(0,120,212);"><b>+27 (21) 715 3171 ext. ' + ext + '</b></span></p>' : '',

    '<div style="line-height:normal;margin:0;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;">',
    '<span style="color:#000;"><b>Cell: &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</b></span>',
    '<span style="color:rgb(0,120,212);"><b>' + cell + '</b></span></div>',

    '<div style="line-height:normal;margin:0;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;">',
    '<span style="color:#000;"><b>Email: &nbsp; &nbsp; &nbsp; </b></span>',
    '<a href="mailto:' + mailtoEmail + '" style="color:rgb(0,120,212);"><b>' + email + '</b></a></div>',

    '<div style="line-height:normal;margin:0;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;">',
    '<span style="color:#000;"><b>Website: </b></span>',
    '<a href="http://www.combustiontechnology.co.za/" style="color:rgb(0,120,212);"><b>www.combustiontechnology.co.za</b></a></div>',

    '<p style="font-size:10pt;"><b><br></b></p>',

    // CTAs
    '<div style="line-height:normal;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;">',
    '<span style="color:#000;"><b>Download our </b></span>',
    '<a href="#" style="color:rgb(200,38,19);"><b>Company Profile</b></a>',
    '<span style="color:#000;"><b>.</b></span></div>',

    '<div style="line-height:normal;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;">',
    '<span style="color:#000;"><b>Please complete our</b></span><b>&nbsp;</b>',
    '<a href="#" style="color:rgb(200,38,19);"><b>Survey</b></a>',
    '<span style="color:#000;"><b>, your opinion matters!</b></span></div>',

    '<div style="line-height:normal;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;">',
    '<a href="#" style="color:rgb(200,38,19);"><b>SIGN UP</b></a>',
    '<span style="color:#000;"><b>&nbsp;for the latest updates on promotions and industry news.</b></span></div>',

    '<div style="font-size:10pt;"><b><br></b></div>',

    // Social icons
    '<div style="line-height:normal;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;">',
    '<a href="#" style="text-decoration:none;"><img alt="LinkedIn" width="38" height="49" src="' + BASE_URL + '/assets/linkedin.png" style="width:38px;height:49px;"></a>',
    '&nbsp;&nbsp;',
    '<a href="#" style="text-decoration:none;"><img alt="Facebook" width="44" height="49" src="' + BASE_URL + '/assets/facebook.png" style="width:44px;height:49px;"></a>',
    '&nbsp;',
    '<a href="#" style="text-decoration:none;"><img alt="YouTube" width="38" height="49" src="' + BASE_URL + '/assets/youtube.png" style="width:38px;height:49px;"></a>',
    '&nbsp; &nbsp;',
    '<a href="#" style="text-decoration:none;"><img alt="Instagram" width="40" height="52" src="' + BASE_URL + '/assets/instagram.png" style="width:40px;height:52px;"></a>',
    '</div>',

    // Banner
    '<div style="font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:12pt;color:#000;">',
    '<img alt="" src="' + BASE_URL + '/assets/banner.png" style="height:auto;max-width:100%;"></div>',

    // POPIA disclaimer
    '<div style="font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:8pt;color:#000;">',
    'At Combustion Technology we take great pride in protecting the personal information that we hold or process about our clients, employees, vendors and other stakeholders.</div>',

    '<div style="line-height:1.2;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:8pt;">',
    '<span style="color:#000;">We are committed to best practice in complying with data protection requirements across our entire global operation.&nbsp;You can read our full POPIA Compliance Statement on our </span>',
    '<a href="#" style="color:rgb(70,120,134);"><u>website</u></a>',
    '<span style="color:#000;">.</span></div>',

    '<div style="line-height:1.2;font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:8pt;">',
    '<span style="color:#000;">Want to change how you receive emails?&nbsp;You can </span>',
    '<a href="#" style="color:rgb(70,120,134);"><u>update your preferences</u></a>',
    '<span style="color:#000;">&nbsp;or send an email to </span>',
    '<a href="mailto:info@combustiontechnology.co.za" style="color:rgb(70,120,134);"><u>info@combustiontechnology.co.za</u></a>',
    '<span style="color:#000;">&nbsp;to be removed from our mailing&nbsp;list.</span></div>',

    '</div>'
  ].join('\n');
}

function esc(s) {
  if (!s) return '';
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
