# Audit des permissions déléguées par les utilisateurs sous Exchange

# Prérequis :

Nécessite Microsoft Exchange Web Services Managed API 1.2 ou versions supérieures
Lien de téléchargement : https://www.microsoft.com/en-us/download/details.aspx?id=42951
Nécessite les Outils d’administration de serveur distant pour Windows 10
Lien de téléchargement : https://www.microsoft.com/fr-FR/download/details.aspx?id=45520
Nécessite d'être membre de Organization Management
Nécessite en outre de posséder le rôle ApplicationImpersonation (présent dans le groupe Hygiene Management)
Nécessite d'être exécuté à partir d'une machine membre du domaine
OutputDir doit se trouver dans le même répertoire que ce script, ainsi que Get-MailboxFolderPermissionEWS.ps1
Ce dernier est à télécharger via ce lien : https://gallery.technet.microsoft.com/scriptcenter/Get-MailboxFolderPermission-fdf1f90f

# Fonctionement
