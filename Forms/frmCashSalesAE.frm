VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCashSalesAE 
   BorderStyle     =   0  'None
   ClientHeight    =   11190
   ClientLeft      =   -30
   ClientTop       =   -405
   ClientWidth     =   20460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCashSalesAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   20460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCashier 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   66
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtTaxBase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12720
      TabIndex        =   65
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Interval        =   360
      Left            =   16200
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   360
      Left            =   15600
      Top             =   720
   End
   Begin VB.CommandButton btnLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   120
      MaskColor       =   &H80000004&
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   10080
      Width           =   6255
   End
   Begin VB.TextBox txtEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   285
      Index           =   9
      Left            =   12720
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   9240
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Ulang Struk"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSettle 
      Caption         =   "Print Struk "
      Height          =   375
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9120
      Width           =   1695
   End
   Begin VB.TextBox txtEntry 
      Height          =   1335
      Index           =   10
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7680
      Width           =   6255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   4455
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorSel    =   1091552
      BackColorBkg    =   -2147483643
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   13080
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   7
      Left            =   12000
      TabIndex        =   7
      Text            =   "0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   9840
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   7680
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   6600
      TabIndex        =   4
      Text            =   "1"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtEntry 
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   4335
   End
   Begin VB.TextBox txtEntry 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Umum"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "FUNGSI AKSI KEYBOARD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   15840
      TabIndex        =   19
      Top             =   1320
      Width           =   4215
      Begin VB.Label Label18 
         Caption         =   "(F12) PRINT ULANG STRUK TRANSAKSI"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "(F11) PRINT STRUK TRANSAKSI"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "(Esc) TUTUP KASIR PENJUALAN"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label14 
         Caption         =   "(F4) BUKA PENJUALAN BARU"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label13 
         Caption         =   "(F5) MELAKUKAN PEMBAYARAN"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label12 
         Caption         =   "(F6) TAMBAH KE KERANJANG BELANJA"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label9 
         Caption         =   "(F7) MENCARI BARANG MANUAL"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "(F8) GUNAKAN BARCODE SCANNER"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Gunakan aksi pada keyboard untuk melakukan proses pelayanan dengan cara menekan tombol berikut ini pada keyboard :"
         Height          =   735
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "(F9) BUKA KASIR PENJUALAN"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   15375
      TabIndex        =   18
      Top             =   600
      Width           =   15375
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   "frmCashSalesAE.frx":6852
      ScaleHeight     =   495
      ScaleWidth      =   20460
      TabIndex        =   14
      Top             =   10695
      Width           =   20460
      Begin VB.Label lblemri 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Harap selalu mengecek kelengkapan pembelian seperti Nama Produk dan Jumlah Pembelian sebelum melakukan input pembayaran."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   120
         Width           =   14055
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Perhatian:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Petugas Kasir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   67
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   2580
      Left            =   15960
      Picture         =   "frmCashSalesAE.frx":B155
      Top             =   6480
      Width           =   3870
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Waktu:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   14760
      TabIndex        =   64
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   16920
      TabIndex        =   63
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblTgl 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18000
      TabIndex        =   62
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblJam 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15600
      TabIndex        =   61
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Info Penggunaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   60
      Top             =   9720
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "Komentar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   120
      TabIndex        =   56
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   12120
      TabIndex        =   55
      Top             =   8880
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   12120
      TabIndex        =   54
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   12120
      TabIndex        =   53
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   12120
      TabIndex        =   52
      Top             =   8160
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   12120
      TabIndex        =   51
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   12120
      TabIndex        =   50
      Top             =   7440
      Width           =   375
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   25
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label lblDiscount 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   49
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblTaxbase 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   46
      Top             =   8160
      Width           =   2415
   End
   Begin VB.Label lblVAT12 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   39
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label lblTot 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   40
      Top             =   8880
      Width           =   2415
   End
   Begin VB.Label Label17 
      Caption         =   "Sub Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   10800
      TabIndex        =   48
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Diskon %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   10800
      TabIndex        =   47
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "PPN 10%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   10800
      TabIndex        =   45
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "Harga Setelah PPN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   9720
      TabIndex        =   44
      Top             =   8520
      Width           =   2055
   End
   Begin VB.Label Label17 
      Caption         =   "Total Pembayaran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   9840
      TabIndex        =   43
      Top             =   8880
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "Jumlah Bayar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   10320
      TabIndex        =   42
      Top             =   9240
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   10320
      TabIndex        =   41
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "Sub Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   13080
      TabIndex        =   38
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Diskon %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   12000
      TabIndex        =   37
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Harga Kotor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9840
      TabIndex        =   36
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "Harga Satuan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7680
      TabIndex        =   34
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   33
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "Deskripsi Produk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   32
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "Barcode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Nama Pembeli"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblChange 
      BackColor       =   &H00000000&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   705
      Left            =   12240
      TabIndex        =   24
      Top             =   9840
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "No.Transaksi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "frmCashSalesAE.frx":C4BC
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblfirma 
      BackStyle       =   0  'Transparent
      Caption         =   "POINT OF SALES POST SHOP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   80
      Width           =   6255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   12000
      Shape           =   4  'Rounded Rectangle
      Top             =   9840
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   0
      Picture         =   "frmCashSalesAE.frx":D11E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20460
   End
End
Attribute VB_Name = "frmCashSalesAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public LLFK                 As Long 'Last Loading FK
Public LLVFK                As Long 'Last Loading Van FK
Public LLDate               As String
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Public blnPaid              As Boolean 'Use to determine if transaction is already paid using frmAmountReceive form

Dim PCase                   As Long 'Pieces per case
Dim PBox                    As Long 'Pieces per box

Dim old_pieces              As Long 'Old pieces value
Dim old_boxes               As Long 'Old boxes value
Dim old_cases               As Long 'Old cases value

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice
Dim bStart                  As Boolean
Dim bScannerOn              As Boolean
Dim blnSave                 As Boolean
Dim intQtyOld               As Integer 'Old txtQty Value. Hold when editing qty

Dim nqty    As Double
Dim nprice  As Double
Dim ndisc   As Double
Dim namt    As Double
Dim ntotal  As Double

Private Sub AddToGrid()
On Error GoTo err
    
    Dim RSStockUnit As New Recordset
    
    Dim intTotalOnhand          As Integer
    Dim intTotalIncoming        As Integer
    Dim intTotalOnhInc          As Integer 'Total of Onhand + Incoming
    Dim intExcessQty            As Integer
    
    Dim intSuggestedQty         As Integer
    Dim blnAddIncoming          As Boolean
    Dim intQtyOrdered           As Integer 'hold the value of txtQty
    Dim intCount                As Integer
    
    If txtEntry(2).Text = "" Then txtEntry(2).SetFocus: Exit Sub
    
    Dim CurrRow As Integer

    Dim intStockID As Integer
    
    CurrRow = getFlexPos(Grid, 8, txtEntry(2).Tag)
    intStockID = txtEntry(2).Tag
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID =" & intStockID & " ORDER BY Stock_Unit.Order ASC", CN, adOpenStatic, adLockOptimistic
    
    If toNumber(txtEntry(5).Text) <= 0 Then
        MsgBox "Harap masukan harga yang valid.", vbExclamation
        txtEntry(5).SetFocus
        Exit Sub
    End If

    intQtyOrdered = txtEntry(4).Text
              
    RSStockUnit.Find "UnitID = " & txtEntry(3).Tag
          
    If RSStockUnit!Onhand < intQtyOrdered Then GoSub GetOnhand

Continue:
    'Save to stock card
    Dim RSStockCard As New Recordset

    RSStockCard.CursorLocation = adUseClient
    RSStockCard.Open "SELECT * FROM Stock_Card", CN, adOpenStatic, adLockOptimistic

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 8) = "" Then
                .TextMatrix(1, 0) = txtEntry(2).Text
                .TextMatrix(1, 1) = txtEntry(3).Text
                .TextMatrix(1, 2) = intQtyOrdered 'txtentry(4).Text
                .TextMatrix(1, 3) = toMoney(txtEntry(5).Text)
                .TextMatrix(1, 4) = toMoney(txtEntry(6).Text)
                .TextMatrix(1, 5) = toMoney(txtEntry(7).Text)
                .TextMatrix(1, 6) = toMoney(txtEntry(8).Text)
                .TextMatrix(1, 7) = txtEntry(3).Tag
                .TextMatrix(1, 8) = intStockID
            Else
AddIncoming:
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = txtEntry(2).Text
                .TextMatrix(.Rows - 1, 1) = txtEntry(3).Text
                .TextMatrix(.Rows - 1, 2) = intQtyOrdered 'txtentry(4).Text
                .TextMatrix(.Rows - 1, 3) = toMoney(txtEntry(5).Text)
                .TextMatrix(.Rows - 1, 4) = toMoney(txtEntry(6).Text)
                .TextMatrix(.Rows - 1, 5) = toMoney(txtEntry(7).Text)
                .TextMatrix(.Rows - 1, 6) = toMoney(txtEntry(8).Text)
                .TextMatrix(.Rows - 1, 7) = txtEntry(3).Tag
                .TextMatrix(.Rows - 1, 8) = intStockID
                
                .FillStyle = 1

                .Row = .Rows - 1
                .Colsel = 6
                If blnAddIncoming = True And intCount = 2 Then
                    .CellForeColor = vbBlue
                    
                    blnAddIncoming = False
                End If
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item produk sudah ada. Apakah anda ingin menambahkannya ?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                'Restore back the invoice amount and discount
                cIGross = cIGross - toNumber(Grid.TextMatrix(.Rowsel, 4))
                lblTotal.Caption = Format$(cIGross, "")
                cIAmount = cIAmount - toNumber(Grid.TextMatrix(.Rowsel, 6))
                lblTot.Caption = Format$(cIAmount, "")
                'Use ExtPrice instead of Sales Price if ExtPrice is more than zero (0)
                cDAmount = cDAmount - toNumber(toNumber(txtEntry(7).Text) / 100) * _
                        (toNumber(toNumber(Grid.TextMatrix(.Rowsel, 2))) * _
                        toNumber(txtEntry(5).Text))
                lblDiscount.Caption = Format$(cDAmount, "")
                
                .TextMatrix(CurrRow, 0) = txtEntry(2).Text
                .TextMatrix(CurrRow, 1) = txtEntry(3).Text
                .TextMatrix(CurrRow, 2) = intQtyOrdered 'txtentry(4).Text
                .TextMatrix(CurrRow, 3) = toMoney(txtEntry(5).Text)
                .TextMatrix(CurrRow, 4) = toMoney(txtEntry(6).Text)
                .TextMatrix(CurrRow, 5) = toMoney(txtEntry(7).Text)
                .TextMatrix(CurrRow, 6) = toMoney(txtEntry(8).Text)
                .TextMatrix(CurrRow, 7) = txtEntry(3).Tag
                .TextMatrix(CurrRow, 8) = intStockID
                                               
                'deduct qty from Stock Unit's table
                RSStockUnit.Filter = "UnitID = " & txtEntry(3).Tag  'getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                
                RSStockUnit!Onhand = RSStockUnit!Onhand + intQtyOld
                
                RSStockUnit.Update
            Else
                Exit Sub
            End If
        End If
        
        RSStockCard.Filter = "StockID = " & intStockID & " AND RefNo2 = '" & txtEntry(1).Text & "'"

        If RSStockCard.RecordCount = 0 Then RSStockCard.AddNew
        
        'Deduct qty solt to stock card
        RSStockCard!Type = "S"
        RSStockCard!UnitID = txtEntry(3).Tag
        RSStockCard!RefNo2 = txtEntry(1).Text
        RSStockCard!Pieces2 = intQtyOrdered
        'Use ExtPrice instead of Sales Price if ExtPrice is more than zero (0)
        RSStockCard!SalesPrice = txtEntry(5).Text
        RSStockCard!StockID = intStockID

        RSStockCard.Update
        
        RSStockUnit.Find "UnitID = " & txtEntry(3).Tag

        'Deduct qty from highest unit breakdown if Onhand is less than qty ordered
        If RSStockUnit!Onhand < intQtyOrdered Then
            DeductOnhand intQtyOrdered, RSStockUnit!Order, True, RSStockUnit
        End If
        
        'deduct qty from Stock Unit's table
        RSStockUnit.Find "UnitID = " & txtEntry(3).Tag
        
        RSStockUnit!Onhand = RSStockUnit!Onhand - intQtyOrdered
        
        RSStockUnit.Update
            
        'Add the amount to current load amount
        cIGross = cIGross + toNumber(txtEntry(6).Text)
        lblTotal.Caption = Format$(cIGross, "")
        cIAmount = cIAmount + toNumber(txtEntry(8).Text)
        'Use ExtPrice instead of Sales Price if ExtPrice is more than zero (0)
        cDAmount = cDAmount + toNumber(toNumber(txtEntry(7).Text) / 100) * _
                (toNumber(intQtyOrdered * _
                toNumber(txtEntry(5).Text)))
        lblDiscount.Caption = Format$(cDAmount, "")
        lblTot.Caption = Format$(cIAmount, "")
        lblTaxbase.Caption = toMoney(lblTotal.Caption * 0.1)
        lblVAT12.Caption = toMoney(lblTotal.Caption - lblTaxbase.Caption)
        txtTaxBase.Text = toMoney(lblTotal.Caption * 0.1)
        'Highlight the current row's column
        .Colsel = 6
        'Display a remove button
        If blnAddIncoming = True Then
            intQtyOrdered = intSuggestedQty
            intCount = 2
            GoSub AddIncoming
            
'            blnAddIncoming = False
        End If
        
        'Reset the entry fields
        ResetEntry
    End With
    
    Exit Sub
    
GetOnhand:

    intTotalOnhand = GetTotalQty("Onhand", RSStockUnit!Order, RSStockUnit!Onhand, RSStockUnit)
    If intTotalOnhand < 0 Then
        MsgBox "Insufficient qty.", vbInformation
    Else
        GoSub Continue
    End If
    
    Exit Sub
    
err:
    Prompt_Err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Function DeductOnhand(QtyNeeded As Integer, ByVal Order As Integer, ByVal blnDeduct As Boolean, rs As Recordset) As Boolean
    Dim Onhand As Boolean
    Dim OrderTemp As Integer
    Dim QtyNeededTemp As Double
    
Reloop:
    OrderTemp = Order
    QtyNeededTemp = QtyNeeded
    rs.Find "Order = " & OrderTemp
    
    
    Do Until Onhand = True 'Or OrderTemp = 1
        If rs!Onhand >= QtyNeededTemp Then
            If blnDeduct = False Then
                DeductOnhand = True
                Exit Function
            Else
                Onhand = True
            End If
            
            If QtyNeededTemp > 0 And QtyNeededTemp < 1 Then
                QtyNeededTemp = 1
            Else
                QtyNeededTemp = CInt(QtyNeededTemp)
            End If
        Else
            OrderTemp = OrderTemp - 1
            If OrderTemp < 1 Then Exit Do
            QtyNeededTemp = (QtyNeededTemp - rs!Onhand) / rs!Qty
            
            rs.MoveFirst
            
            rs.Find "Order = " & OrderTemp
        End If
    Loop
    
    If Onhand = True Then
        Do
            rs!Onhand = rs!Onhand - QtyNeededTemp
            OrderTemp = OrderTemp + 1
            
            rs.MoveFirst
            rs.Find "Order = " & OrderTemp
            
            rs!Onhand = rs!Onhand + (QtyNeededTemp * rs!Qty)
            
            rs.Update
            
            Onhand = False
            
            If OrderTemp = Order Then
                DeductOnhand = True
                Exit Do
            Else
                GoSub Reloop
            End If
        Loop
    Else
        DeductOnhand = False
    End If
End Function

'Get the total Qty onhand, incoming and total of onhand and incoming
Private Function GetTotalQty(strField As String, Order As Integer, intOnhand As Integer, rs As Recordset) As Integer
    Dim strFieldValue As Integer
    Dim intOrder As Integer
    
    GetTotalQty = intOnhand
    
    intOrder = Order - 1
    
    Do Until intOrder < 1
        rs.MoveFirst
        rs.Find "Order = " & intOrder
        
        If strField = "Onhand" Then
            strFieldValue = rs!Onhand
        ElseIf strField = "Incoming" Then
            strFieldValue = rs!Incoming
        Else
            strFieldValue = rs!TotalQty
        End If
        
        GetTotalQty = GetTotalQty + GetTotalUnitQty(Order, intOrder, strFieldValue, rs)
        intOrder = intOrder - 1
    Loop
End Function

'This function is called by GetTotalQty Function
Private Function GetTotalUnitQty(Order As Integer, ByVal Ordertmp As Integer, intOnhand As Integer, rs As Recordset)
    GetTotalUnitQty = 1
    Do Until Order = Ordertmp
        Ordertmp = Ordertmp + 1
        
        rs.MoveNext
        
        GetTotalUnitQty = GetTotalUnitQty * rs!Qty
    Loop
    GetTotalUnitQty = intOnhand * GetTotalUnitQty
End Function

Private Function GetIncoming(QtyNeeded As Integer, ByVal Order As Integer, ByVal blnDeduct As Boolean, rs As Recordset) As Boolean
    Dim Onhand As Boolean
    Dim OrderTemp As Integer
    Dim QtyNeededTemp As Double
    
Reloop:
    OrderTemp = Order
    QtyNeededTemp = QtyNeeded
    rs.Find "Order = " & OrderTemp
    
    
    Do Until Onhand = True 'Or OrderTemp = 1
        If rs!Incoming >= QtyNeededTemp Then
            If blnDeduct = False Then
                GetIncoming = True
                Exit Function
            Else
                Onhand = True
            End If
            
            If QtyNeededTemp > 0 And QtyNeededTemp < 1 Then
                QtyNeededTemp = 1
            Else
                QtyNeededTemp = CInt(QtyNeededTemp)
            End If
        Else
            OrderTemp = OrderTemp - 1
            If OrderTemp < 1 Then Exit Do
            QtyNeededTemp = (QtyNeededTemp - rs!Incoming) / rs!Qty
            
            rs.MoveFirst
            
            rs.Find "Order = " & OrderTemp
        End If
    Loop
    
    If Onhand = True Then
        Do
            rs!Incoming = rs!Incoming - QtyNeededTemp
            OrderTemp = OrderTemp + 1
            
            rs.MoveFirst
            rs.Find "Order = " & OrderTemp
            
            rs!Incoming = rs!Incoming + (QtyNeededTemp * rs!Qty)
            
            rs.Update
            
            Onhand = False
            
            If OrderTemp = Order Then
                GetIncoming = True
                Exit Do
            Else
                GoSub Reloop
            End If
        Loop
    Else
        GetIncoming = False
    End If
End Function

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.Rowsel)
    End With
End Sub

Private Function GetUnitID(ByVal sUnit As String) As Long
  Dim rs As New ADODB.Recordset
  Dim SQL As String
  
  
  SQL = "SELECT Unit.unit_id From unit WHERE (((Unit.unit)='" & Replace(sUnit, "'", "''") & "'))"
  
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    GetUnitID = rs!unit_id
  Else
    GetUnitID = 0
  End If
  
  
  rs.Close
  Set rs = Nothing
End Function

Private Function getunitid1(ByVal Stock As String) As Long
  Dim rs As New ADODB.Recordset
  Dim SQL As String
  
  
  SQL = "SELECT Unit.unit_id " _
  & "FROM Stocks LEFT JOIN Unit ON Stocks.unit_id = Unit.unit_id " _
  & "WHERE (((Stocks.stock)='" & Replace(Stock, "'", "''") & "'))"

  
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    getunitid1 = IIf(IsNull(rs!unit_id), 5, rs!unit_id)
  Else
    getunitid1 = 0
  End If
  
  
  rs.Close
  Set rs = Nothing
End Function


Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tUser1 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: n/a" & vbCrLf & _
           "Modified By: n/a", vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tUser1 = vbNullString
End Sub



Private Sub cmdPrint_Click()
    If blnSave = False Then
        btnLabel.Caption = "Selesaikan transaksi sebelum melakukan print struk pembelian."
        Exit Sub
    End If
    
    PrintInvoice txtEntry(0).Text
End Sub

Private Sub PrintInvoice(InvoiceNo As String)
    With frmReports
        .strReport = "Receipt"
        .strWhere = "{Cash_Sales.InvoiceNo} ='" & InvoiceNo & "'"
        
        frmReports.show vbModal
    End With
End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Function makecashslip() As String
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  Dim Temp As String
  
  SQL = "SELECT Last(Cash_Sales.cash_sales_id) AS LastOfcash_sales_id " _
  & "From Cash_Sales " _
  & "ORDER BY Last(Cash_Sales.cash_sales_id)"
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    Temp = IIf(IsNull(rs!LastOfcash_sales_id), 0, rs!LastOfcash_sales_id) + 1
  Else
    Temp = 1
  End If

  makecashslip = Format(Temp, "0000000000")
  rs.Close
  Set rs = Nothing
End Function

Private Sub cmdSettle_Click()
    Dim rs As New Recordset
        
    If blnPaid = False Then
        btnLabel.Caption = "Transaksi belum dibayar."
        Exit Sub
    End If
    
    rs.Open "SELECT * FROM Cash_Sales WHERE CashSalesID=" & PK, CN, adOpenStatic, adLockOptimistic

    'Verify the entries
    If txtEntry(0).Text = "" Then
        MsgBox "Harap masukan slip transaksi.", vbExclamation
        txtEntry(0).SetFocus
        Exit Sub
    End If
   
    If cIRowCount < 1 Then
        MsgBox "Harap masukan produk pada transaksi sebelum transaksi disimpan pada database.", vbExclamation
        txtEntry(2).SetFocus
        Exit Sub
    End If
    
    If isRecordExist("Cash_Sales", "InvoiceNo", txtEntry(0).Text, True) = True Then
'        MsgBox "Cash slip already exist. Please change it.", vbExclamation
        btnLabel.Caption = "Transaksi sudah disimpan."
        
        txtEntry(0).SetFocus
        Exit Sub
    End If

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Cash_Sales_Detail WHERE CashSalesID=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    'Save the record
    With rs
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![CashSalesID] = PK
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        Else
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If
        ![InvoiceNo] = txtEntry(0).Text
        ![Date] = Now()
        ![SoldTo] = txtEntry(1).Text
        ![Gross] = toNumber(lblTotal.Caption)
        ![Discount] = lblDiscount.Caption
        ![TaxBase] = lblTaxbase.Caption
        ![Vat] = lblVAT12.Caption
        ![NetAmount] = lblTot.Caption
        ![Tendered] = txtEntry(9).Text
        ![Change] = lblChange.Caption
        ![Remarks] = txtEntry(10).Text
        ![Cashier] = txtCashier.Text

        .Update

    End With

    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then

                RSDetails.AddNew

                'RSDetails![PK] = getIndex("tbl_AR_InvoiceDetails")

                RSDetails![CashSalesID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 8))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 2))
                RSDetails![UnitID] = toNumber(.TextMatrix(c, 7))
                RSDetails![Price] = toNumber(.TextMatrix(c, 3))
                RSDetails![Discount] = toNumber(.TextMatrix(c, 5))
                
                RSDetails.Update
            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set RSDetails = Nothing

    CN.CommitTrans

    blnSave = True

    HaveAction = True
    Screen.MousePointer = vbDefault

    btnLabel.Caption = "Changes in record has been successfully saved."
    
    PrintInvoice txtEntry(0).Text
    
    Exit Sub
err:
    blnSave = False

'    CN.RollbackTrans
    Prompt_Err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Function getcashsalesid() As Long
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  
  SQL = "SELECT Cash_Sales.cash_sales_id " _
  & "From Cash_Sales " _
  & "WHERE (((Cash_Sales.cash_slip)='" & Replace(txtEntry(0).Text, "'", "''") & "') AND " _
  & "((Cash_Sales.sold_to)='" & Replace(txtEntry(1).Text, "'", "''") & "'))"
  
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    getcashsalesid = rs!Cash_Sales_ID
  Else
    getcashsalesid = 0
  End If
  
  rs.Close
  Set rs = Nothing
End Function

Private Function GetStockID(ByVal Stock As String) As Long
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  
  SQL = "SELECT Stocks.stock_id " _
  & "From Stocks " _
  & "WHERE (((Stocks.stock)='" & Replace(Stock, "'", "''") & "'))"
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    GetStockID = rs!stock_id
  Else
    GetStockID = 0
  End If
  
  rs.Close
  Set rs = Nothing
End Function

Private Function saveslipdetail(ByVal b As Boolean) As Boolean
On Error GoTo errhandler
  Dim i As Long
  Dim SQL As String
  
  
  With Grid
    For i = 1 To .Rows - 1
      If b Then
        SQL = "insert into cash_sales_detail(cash_sales_id, stock_id, qty, unit, price, " _
        & "amount_gross, discount_percent, discount_amount, amount_net, net_price, [date]) " _
        & "values(" & getcashsalesid & ", " _
        & GetStockID(.TextMatrix(i, 1)) & ", " _
        & .TextMatrix(i, 2) & ", " _
        & getunitid1(.TextMatrix(i, 1)) & ", " _
        & CDbl(.TextMatrix(i, 3)) & ", " _
        & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) & ", " _
        & .TextMatrix(i, 4) & ", " _
        & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) * (.TextMatrix(i, 4) / 100) & ", " _
        & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) - (CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) * (.TextMatrix(i, 4) / 100)) & ", " _
        & "0, #" _
        & Format(Date, "mm/dd/yy") & "#)"
      Else
        'sql = "update cash_sales_detail set " _
        & "cash_sales_id=" & getcashsalesid & ", " _
        & "stock_id=" & getstockid(.TextMatrix(i, 1)) & ", " _
        & "qty=" & .TextMatrix(i, 2) & ", " _
        & "unit=0, " _
        & "price=" & CDbl(.TextMatrix(i, 3)) & ", " _
        & "amount_gross=" & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) & ", " _
        & "discount_percent=" & .TextMatrix(i, 4) & ", " _
        & "discount_amount=" & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) * (.TextMatrix(i, 4) / 100) & ", " _
        & "amount_net=" & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) - (CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) * (.TextMatrix(i, 4) / 100)) & ", " _
        & "[date]=#" & Format(Date, "mm/dd/yy") & "#"
      End If
      CN.Execute SQL
    Next
  End With
  saveslipdetail = True
  
  Exit Function
errhandler:
  MsgBox "Error: " & err.Description & vbCr _
  & "Form: frmCashSalesAE" & vbCr _
  & "Function: saveslipdetail", vbExclamation
  saveslipdetail = False
End Function

Private Function saveslip(ByVal b As Boolean) As Boolean
On Error GoTo errhandler
  Dim SQL As String
  
  If b Then
    SQL = "insert into cash_sales(cash_slip, [date], sold_to, payment_type, gross, " _
    & "discount, tax_base, vat, amount_net, cash, remarks) " _
    & "values('" & Replace(txtEntry(0).Text, "'", "''") & "', #" _
    & Format(Date & " " & Time, "mm/dd/yy hh:mm:ss") & "#, '" _
    & Replace(txtEntry(1).Text, "'", "''") & "', '', " _
    & CDbl(lblTotal.Caption) & ", 0, 0, " _
    & CDbl(lblVAT12.Caption) & ", " _
    & CDbl(lblTot.Caption) & ", " _
    & CDbl(txtEntry(8).Text) & ", '" _
    & Replace(txtEntry(9).Text, "'", "''") & "')"
  Else
    'sql = "update cash_sales set " _
    & "cash_slip='" & Replace(txtEntry(0).Text, "'", "''") & "', " _
    & "[date]=#" & Format(Date, "mm/dd/yy") & "#, " _
    & "sold_to='" & Replace(txtEntry(1).Text, "'", "''") & "', " _
    & "payment_type='', " _
    & "gross=" & CDbl(lblTotal.Caption) & ", " _
    & "discount=0, " _
    & "tax_base=0, " _
    & "vat=" & CDbl(lblVAT12.Caption) & ", " _
    & "amount_net=" & CDbl(lblTot.Caption) & ", " _
    & "cash=" & CDbl(txtEntry(8).Text) & ", '" _
    & "remarks='" & Replace(txtEntry(9).Text, "'", "''") & "'"
  End If
  
  CN.Execute SQL
  saveslip = True
  
  Exit Function
  
errhandler:
  MsgBox "Error: " & err.Description & vbCr _
  & "Form: frmCashSalesAE" & vbCr _
  & "Function: saveslip", vbExclamation
  saveslip = False
End Function

Private Sub cmdSettle_KeyDown(KeyCode As Integer, Shift As Integer)
'  Form_KeyDown KeyCode, Shift
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then
        Unload Me
    End If
End Sub

Private Sub addtolist()
  
  With Grid
    If .Rows = 2 And .TextMatrix(1, 0) = "" Then
      .TextMatrix(1, 0) = txtEntry(2).Text
      .TextMatrix(1, 1) = txtEntry(3).Text
      .TextMatrix(1, 2) = txtEntry(4).Text
      .TextMatrix(1, 3) = Format(txtEntry(5).Text, "#,##0.00")
      .TextMatrix(1, 4) = txtEntry(6).Text
      .TextMatrix(1, 5) = Format(txtEntry(7).Text, "#,##0.00")
    Else
      .Rows = .Rows + 1
      .TextMatrix(.Rows - 1, 0) = txtEntry(2).Text
      .TextMatrix(.Rows - 1, 1) = txtEntry(3).Text
      .TextMatrix(.Rows - 1, 2) = txtEntry(4).Text
      .TextMatrix(.Rows - 1, 3) = Format(txtEntry(5).Text, "#,##0.00")
      .TextMatrix(.Rows - 1, 4) = txtEntry(6).Text
      .TextMatrix(.Rows - 1, 5) = Format(txtEntry(7).Text, "#,##0.00")
    End If
  End With
  txtEntry(2).Text = ""
  txtEntry(3).Text = ""
  txtEntry(4).Text = "1"
  txtEntry(5).Text = "0.00"
  txtEntry(6).Text = "0"
  txtEntry(7).Text = "0.00"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If State = adStateEditMode Or State = adStateAddMode Then Exit Sub
    
    Select Case KeyCode
      Case vbKeyEscape
        blnSave = False
        
        lblclose_Click
      
      Case vbKeyF4
        Form_KeyDown vbKeyF9, Shift
      Case vbKeyF5
        frmAmountReceive.txtAmt.Text = Format(lblTot.Caption, "")
        frmAmountReceive.show 1
      Case vbKeyF6

        txtEntry_KeyPress 7, 13
        
      Case vbKeyF7        'lookup table
        If Not bStart Then
          btnLabel.Caption = "Untuk memulai transaksi silahkan tekan tombol (F9) pada keyboard."
          Exit Sub
        End If
        frmLookup.show 1
        txtEntry(2).SetFocus
      Case vbKeyF8        'ready scanner for scanning
        If Not bStart Then
          MsgBox "Untuk memulai transaksi silahkan tekan tombol (F9) pada keyboard.", vbInformation
          Exit Sub
        End If
        
        txtEntry(2).SetFocus
      Case vbKeyF9
        bStart = True
        blnSave = False
        blnPaid = False
        
        txtEntry(0).Text = ""
        txtEntry(1).Text = ""
        
        ResetEntry
        
        txtEntry(0).SetFocus
        ntotal = 0
        
        InitGrid
        
        CN.BeginTrans
        
        GeneratePK
        
        txtEntry(0).Text = Format(PK, "0000000000")
        
        txtEntry(1).SetFocus
        'Menampilkan nama kasir
        txtCashier.Text = CurrUser.USER_NAME
        lblTotal.Caption = "0.00"
        lblDiscount.Caption = "0.00"
        lblTaxbase.Caption = "0.00"
        lblVAT12.Caption = "0.00"
        lblTot.Caption = "0.00"
        lblChange.Caption = "0.00"
        
      Case vbKeyF11
        cmdSettle_Click
      Case vbKeyF12
        cmdPrint_Click
      
      
      Case vbKeyEnd
      Case vbKeyHome
      Case vbKeyUp, vbKeyPageUp
        'If Shift = vbCtrlMask Then
        '  cmdFirst_Click
        'Else
        '  cmdPrevious_Click
        'End If
      Case vbKeyDown, vbKeyPageDown
        'If Shift = vbCtrlMask Then
        '  cmdLast_Click
        'Else
        '  cmdNext_Click
        'End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub


'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Cash_Sales")
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        '.FixedCols = 1
        .Cols = 9 '13
        .Colsel = 7
        'Initialize the column size
        .ColWidth(0) = 1455
        .ColWidth(1) = 5460
        .ColWidth(2) = 660
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1545
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = "Barcode"
        .TextMatrix(0, 1) = "Deskripsi Produk"
        .TextMatrix(0, 2) = "Qty"
        .TextMatrix(0, 3) = "Harga Satuan"
        .TextMatrix(0, 4) = "Harga Kotor"
        .TextMatrix(0, 5) = "Diskon"
        .TextMatrix(0, 6) = "Total Harga"
        .TextMatrix(0, 7) = "Unit ID"
        .TextMatrix(0, 8) = "Stock ID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbRightJustify
        .ColAlignment(3) = vbRightJustify
        .ColAlignment(4) = vbRightJustify
        .ColAlignment(5) = vbRightJustify
        .ColAlignment(6) = vbRightJustify
    End With
End Sub

Private Sub ResetEntry()
    txtEntry(2).Text = ""
    txtEntry(3).Text = ""
    txtEntry(4).Text = "1"
    txtEntry(5).Text = "0.00"
    txtEntry(6).Text = "0"
    txtEntry(7).Text = "0.00"
    txtEntry(8).Text = "0.00"
    txtEntry(9).Text = "0.00"
'    lblTotal.Caption = "0.00"
'    lblDiscount.Caption = "0.00"
'    lblTaxbase.Caption = "0.00"
'    lblVAT12.Caption = "0.00"
'    lblTot.Caption = "0.00"
'    lblChange.Caption = "0.00"
    
'    SendKeys ("{tab}")
'    SendKeys ("{tab}")

    txtEntry(2).SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmCashSalesAE = Nothing
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub lblclose_Click()
On Error Resume Next

    If blnSave = False Then CN.RollbackTrans
    Unload Me
End Sub

'Private Sub lblclose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'imgcloseactive.Visible = True
'End Sub

Private Sub Timer1_Timer()
lblJam.Caption = Time
End Sub

Private Sub Timer2_Timer()
lblTgl.Caption = Format(Date, "dd MMMM YYYY")
End Sub

Private Sub txtEntry_Change(Index As Integer)
    If Index = 2 And Len(txtEntry(2).Text) = 13 Then
        rs.Open "SELECT StockID, Barcode, Stock, SalesPrice, UnitID FROM qry_Stock_Unit WHERE Barcode = '" & txtEntry(2).Text & "' ORDER BY [Order] DESC", CN, adOpenStatic, adLockOptimistic

        With rs
            If .RecordCount = 0 Then
                MsgBox "No available stock.", vbInformation
                
                .Close
                Exit Sub
            End If
            
            txtEntry(2).Tag = !StockID
            txtEntry(3).Tag = !UnitID 'Add Unit ID to Stock control to avoid adding of dummy control
            txtEntry(3).Text = !Stock
            txtEntry(5).Text = toMoney(!SalesPrice)
            
            .Close
        End With
    End If
    
    If Index = 4 Or Index = 5 Or Index = 7 Then
        txtEntry(6).Text = toMoney((toNumber(txtEntry(4).Text) * toNumber(txtEntry(5).Text)))
        txtEntry(8).Text = toMoney((toNumber(txtEntry(4).Text) * _
                toNumber(txtEntry(5).Text)) - _
                ((toNumber(txtEntry(7).Text) / 100) * _
                toNumber(toNumber(txtEntry(4).Text) * _
                toNumber(txtEntry(5).Text))))
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
    If Index = 8 Then
'        cmdSave.Default = False
    End If
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index = 7 Then
        AddToGrid
    End If
    
    If Index > 1 And Index < 7 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index = 5 Then
        txtEntry(Index).Text = toMoney(txtEntry(Index).Text)
    End If
End Sub

Private Sub txtTaxBase_Change()
    lblVAT12.Caption = toMoney(toNumber(txtTaxBase.Text) + toNumber(lblTotal.Caption))
    lblTot.Caption = toMoney(lblVAT12.Caption)
End Sub
