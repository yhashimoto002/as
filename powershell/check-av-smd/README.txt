
������

-- check-av-smd-once.ps1

���[�U�� av-smd.bin �� av-smd.bin.sig2 �̃^�C���X�^���v���`�F�b�N���܂��B

�X�N���v�g�Ɠ����t�H���_�� result_AV_yyyyMMdd.csv �Ƃ������O�Ń`�F�b�N���ʂ�
�o�͂���Ă����܂��B

"User"	"File"	"TimeStamp"	"CheckDate"	"Result"
"adachicity"	"av-smd.bin"	"2018/07/05 6:13:36"	"2018/07/05 19:00:11"	"OK"
"adachicity"	"av-smd.bin.sig2"	"2018/07/05 6:13:43"	"2018/07/05 19:00:12"	"OK"
"AioiCity"	"av-smd.bin"	"2018/07/05 6:13:36"	"2018/07/05 19:00:12"	"OK"
"AioiCity"	"av-smd.bin.sig2"	"2018/07/05 6:13:43"	"2018/07/05 19:00:13"	"OK"
(...)

�w��������ߋ��̃^�C���X�^���v�����o���ꂽ�� maillist_NG.txt �ɋL�ڂ��ꂽ
���[���A�h���X���ĂɃ��[�����ʒm����܂��B

����ɁA���ʂ� OK �̏ꍇ�ł����T���j���� 8:00 �ɒʒm���܂��B
����ʒm����j���Ǝ��Ԃ͕����w��\�ł�

-- check-all-once.ps1

���[�U�̂��ׂĂ� bin �t�@�C���̃^�C���X�^���v���擾���邾���ł��B
�{�X�N���v�g�ł͓��ɒʒm�͂��܂���B

�X�N���v�g�Ɠ����t�H���_�� result_ALL_yyyyMMdd.csv �Ƃ������O�Ń`�F�b�N���ʂ�
�o�͂���Ă����܂��B

"User"	"Station"	"File"	"TimeStamp"	"CheckDate"
"adachicity"	"SDSR"	"license.bin"	"2016/12/22 16:45:38"	"2018/07/06 11:20:55"
"adachicity"	"SDSR"	"license.bin.sig2"	"2016/12/22 16:45:39"	"2018/07/06 11:20:56"
"adachicity"	"SDSR"	"config.bin"	"2017/03/02 10:20:36"	"2018/07/06 11:20:56"
"adachicity"	"SDSR"	"config.bin.sig2"	"2017/03/02 10:20:36"	"2018/07/06 11:20:57"
(...)


���g����

-- check-av-smd-once.ps1

1. �C�ӂ̃t�H���_�Ɉȉ��̂悤�Ƀt�@�C���ƃt�H���_��z�u���܂�

	\- check-av-smd-once.ps1 �t�@�C��
	\- settings.ini �t�@�C��
	\- user.txt �t�@�C��
	\- general �t�H���_
		\- Invoke-WebrequestToUpdateServer.ps1
		\- Send-MailMessage-Net.ps1

2. settings.ini ���e�L�X�g�G�f�B�^�ŊJ���A�K�v�ɉ����Đݒ��ύX���܂�
�@ ���[���̒ʒm�� (mailToInNG�AmailToInOK) �ɂ͒��ӂ��Ă�������

3. 1�񂾂����s����ꍇ�� PowerShell �𗧂��グ�ăX�N���v�g�����s���܂��B

	PS> .\check-av-smd-once.ps1

4. ������s����ꍇ�� Windows �̃^�X�N�X�P�W���[���ɃX�N���v�g��o�^���܂�

	�V�K�Ń^�X�N��o�^����ɂ� �^�X�N�X�P�W���[�� > �^�X�N�̍쐬 ���J���A
	�K�v�Ȑݒ��o�^���܂��B

	1���Ԓu���Ɏ��s��������ꍇ�� �g���K�[ �� ���� �̐ݒ��̃X�N���[���V���b�g��
	taskscheduler_sample01.png �� taskscheduler_sample02.png �ɎB���Ă��܂��̂ŁA
	�Q�l�ɂ��Ă��������B
	
	�������� �^�X�N�̃C���|�[�g ��� AV signature update check.xml ���w�肵��
	�C���|�[�g���܂��B
	���̏ꍇ�́AC:\work\check-av-smd �t�H���_���쐬���āA�����ɃX�N���v�g����
	�z�u���Ă��������B

-- check-all-once.ps1

1. �C�ӂ̃t�H���_�Ɉȉ��̂悤�Ƀt�@�C���ƃt�H���_��z�u���܂�

	\- check-all-once.ps1 �t�@�C��
	\- user.txt �t�@�C��
	\- general �t�H���_
		\- Invoke-WebrequestToUpdateServer.ps1
		\- Send-MailMessage-Net.ps1

2. PowerShell �𗧂��グ�ăX�N���v�g�����s���܂��B

	PS> .\check-all-once.ps1


�� ����

�EPowerShell 4.0 �������Ɛ��������삵�܂���BWindows 7 ���ƃo�[�W�����A�b�v���Ă��Ȃ���� 2.0 �̂܂܂ł��B

PowerShell ���J���Ĉȉ��̃R�}���h�����s���āu2�v���\�������ꍇ�APowerShell 4.0 �ȏ��
�C���X�g�[�����邩�AWindows Server 2012 R2 �Ȃǂ̊��Ŏ��s���Ă��������B

PS> $PSVersionTable.PSVersion.Major

PowerShell �̃A�b�v�O���[�h���@�͈ȉ��� URL ���Q�l�ɂȂ�܂��B

Windows PowerShell �̃C���X�g�[�� - ������ Windows PowerShell ���A�b�v�O���[�h����
https://docs.microsoft.com/ja-jp/powershell/scripting/setup/installing-windows-powershell?view=powershell-6#upgrading-existing-windows-powershell


�E�X�N���v�g�̎��s�|���V�[�� Restricted �̏ꍇ�́A�X�N���v�g�����s�ł��܂���B
���̏ꍇ�́A�ȉ��̃R�}���h�� RemoteSigned �ɕύX���Ă�����s���Ă��������B

> Get-ExecutionPolicy
Restricted

> Set-ExecutionPolicy RemoteSigned

> Get-ExecutionPolicy
RemoteSigned

�E���s�������񋁂߂���ꍇ�͈ȉ��̃R�}���h�����s���Ă��������B

> Unblock-File .\*.ps1


������
2018/6/8 ���{ �V�K�쐬
2018/7/6 ���{ ���j���[�A��


