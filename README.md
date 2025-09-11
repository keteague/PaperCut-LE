# OVERVIEW
Leverage the use of Posh-ACME to issue a Let's Encrypt TLS certificate for use with PaperCut MF.

# PREREQUISITES
  * PowerShell 7
  * OpenSSL (Lite) - https://slproweb.com/products/Win32OpenSSL.html

# INSTALL
1. On your computer that's running the PaperCut services, download the driver script and its related modules.  Note that modules should be in a "Modules" sub-directory from where the driver script (Setup-PaperCutLetsEncrypt.ps1) runs.  In my test environment, I dropped them into C:\Scripts.
2. Edit Setup-PaperCutLetsEncrypt.ps1 and modify the $Fqdn and $ContactEmail variables
3. Open PowerShell 7 as Administrator
4. Run Setup-PaperCutLetsEncrypt.ps1

This will:
  * Request a certificate from Let's Encrypt
  * Convert it to .PFX format
  * Import the certificate into PaperCut
  * Setup a Task Scheduler event to run daily to check if renewal is necessary.

The renewal task will check the expiration date of the certificate and only attempt to request a new certificate when it's 30 days from expiring.
