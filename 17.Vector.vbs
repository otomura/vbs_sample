Option Explicit

Class Vector
	
	'�f�[�^(���I�z��)
	Private m_data()
	'�f�[�^��
	Private m_dataCount
	
'Public
	
	' �v�f�ǉ�
	Public Sub Add(value)
		If Ubound(m_data) < m_dataCount + 1 Then
			Call ExpandData()
		End If
		Call Bind(m_data(m_dataCount), value)
		m_dataCount = m_dataCount + 1
	End Sub
	
	' �v�f�擾
	' Default �ɂ��邱�ƂŁA"�I�u�W�F�N�g��(�v�f�ԍ�)"�ŃA�N�Z�X�ł���
	Public Default Property Get Item(num)
		Call Bind(Item, m_data(num))
	End Property

	' �v�f���擾
	Public Property Get Count
		Count = m_dataCount
	End Property
	
	' �C�e���[�^�擾
	Public Function Iterator()
		Dim ite : Set ite = New VectorIterator
		Set ite.Taget = Me
		Set Iterator = ite
	End Function
	
	' �f�o�b�O�p �����񉻃��\�b�h
	Public Function ToString()
		Dim result
		Dim data
		For Each data in m_data
			If Not IsObject(data) Then
				result = result & data & " "
			Else
				'�I�u�W�F�N�g�̏ꍇ�͌^�\���̂�.
				result = result & TypeName(data) & " "
			End If
		Next
		ToString = result
	End Function
	
'Private

	'�T�C�Y�g��
	Private Sub ExpandData
		Redim Preserve m_data(Ubound(m_data)*2 + 1)
	End Sub
	
	'�R���X�g���N�^
	Private Sub Class_Initialize
		m_dataCount = 0
		Redim m_data(0)
	End Sub
	
	'�I�u�W�F�N�g�ݒ�p
	Private Sub Bind(var, val)
		If IsObject(val) = True then
			Set var = val
		Else
			var = val
		End If
	End Sub
End Class

'�C�e���[�^
Class VectorIterator

	'�Q�Ƃ���Vector
	Private m_vector

	'Next �ŕԂ��l�̃C���f�b�N�X
	Private m_NextIndex
	
'Public

	'�擾�ł���v�f���邩
	Public Function HasNext()
		hasNext = (m_NextIndex < m_vector.Count)
	End Function
	
	'�v�f�擾
	'Next �͗\��ꂾ���炩�R���p�C���G���[�ɂȂ�̂� GetNext
	Public Function GetNext()
		Call Bind(GetNext, m_vector(m_NextIndex))
		m_NextIndex = m_NextIndex + 1
	End Function
	
	' m_vector �̃v���p�e�B�ݒ�
	' �R���X�g���N�^�����Ȃ��̂�..
	Public Property Set Taget(vector)
		Set m_vector = vector
	End Property
	
'Private	
	'�R���X�g���N�^
	Private Sub VectorIterator_Initialize
		m_NextIndex = 0
	End Sub
	
	'�I�u�W�F�N�g�ݒ�p
	Private Sub Bind(var, val)
		If IsObject(val) = True then
			Set var = val
		Else
			var = val
		End If
	End Sub
End Class
