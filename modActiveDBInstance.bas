Attribute VB_Name = "modActiveDBInstance"
Option Explicit
Private Const MOD_NAME = "modActiveDBInstance"

'This instance of the database MAY NOT be accessed from modDB or modCommon.
'This is meant to be the active database, not some one-and-only global copy.
Public ActiveDBInstance As EJTSClientsDB

