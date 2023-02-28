Class {
	#name : #GtExcelFile,
	#superclass : #Object,
	#instVars : [
		'sheets',
		'filename'
	],
	#classInstVars : [
		'generatorProgram'
	],
	#category : #GToolkit4Excel
}

{ #category : #accessing }
GtExcelFile class >> buildProgram [
	| process stdout stderr error |
	process := (GtExternalProcessBuilder new: 'dotnet')
			arg: 'build';
			arg: '--configuration';
			arg: 'Release';
			arg: self generatorDirectory fileReference pathString;
			pipeStdout;
			pipeStderr;
			spawn.
	stdout := process asynchronousStdout.
	stderr := process asynchronousStderr.
	process wait isSuccess
		ifFalse: [ error := stderr pollString.
			error isEmpty ifTrue: [ error := stdout pollString ].
			self error: error ].
	^ stdout pollString
]

{ #category : #accessing }
GtExcelFile class >> generatorDirectory [
	^ GtResourceReference
		path: Path * 'feenkcom' / 'gt4excel' / 'dotnet' / 'ExcelGenerator'
]

{ #category : #accessing }
GtExcelFile class >> generatorProgram [
	^ generatorProgram
		ifNil: [ self buildProgram.
			(self generatorDirectory / 'bin' / 'Release') directories
				do: [ :each | 
					(each filesMatching: 'ExcelGenerator*')
						do: [ :file | 
							| filename |
							filename := file.
							[ filename basename includes: $. ]
								whileTrue: [ filename := filename withoutExtension ].
							^ generatorProgram := filename pathString ] ].
			'ExcelGenerator' ]
]

{ #category : #accessing }
GtExcelFile class >> generatorProgram: aString [
	generatorProgram := aString
]

{ #category : #accessing }
GtExcelFile class >> outputCommand: aString withArgs: anArray on: aStream [
	aStream
		nextPutAll: aString;
		nextPut: $(.
	anArray
		do: [ :each | aStream nextPutAll: (self stringFor: each) ]
		separatedBy: [ aStream nextPutAll: ', ' ].
	aStream
		nextPut: $);
		nextPutAll: String crlf
]

{ #category : #private }
GtExcelFile class >> stringFor: each [
	each isString
		ifTrue: [ ^ '"'
				, ((((each copyReplaceAll: '"' with: '""') copyReplaceAll: '\' with: '\\')
						copyReplaceAll: String cr
						with: '\r') copyReplaceAll: String lf with: '\n') , '"' ].
	^ each asString
]

{ #category : #accessing }
GtExcelFile >> filename [
	^ filename
]

{ #category : #accessing }
GtExcelFile >> filename: aString [
	filename := aString asFileReference fullName
]

{ #category : #accessing }
GtExcelFile >> gtActionSaveFor: anAction [
	<gtAction>
	^ anAction button
		icon: BrGlamorousVectorIcons folder;
		tooltip: 'Save';
		action: [ :button | 
			self save.
			button phlow spawnObject: self filename asFileReference ]
]

{ #category : #views }
GtExcelFile >> gtSheetsFor: aView [
	<gtView>
	^ aView list
		title: 'Sheets';
		items: [ sheets ]
]

{ #category : #testing }
GtExcelFile >> hasSheetNamed: aString [
	^ sheets anySatisfy: [ :each | each name = aString ]
]

{ #category : #accessing }
GtExcelFile >> initialize [
	super initialize.
	sheets := OrderedCollection new
]

{ #category : #accessing }
GtExcelFile >> isEmpty [
	^ sheets isEmpty
]

{ #category : #saving }
GtExcelFile >> save [
	| process stdout stderr error |
	process := (GtExternalProcessBuilder new: self class generatorProgram)
			pipeStdin;
			pipeStdout;
			pipeStderr;
			spawn.
	stdout := process asynchronousStdout.
	stderr := process asynchronousStderr.
	self saveOn: process stdin.
	process wait isSuccess
		ifFalse: [ error := stderr pollString.
			error isEmpty ifTrue: [ error := stdout pollString ].
			self error: error ]
]

{ #category : #saving }
GtExcelFile >> saveOn: aStream [
	self class
		outputCommand: 'CreateDocument'
		withArgs: {filename}
		on: aStream.
	sheets do: [ :each | each saveOn: aStream ].
	self class
		outputCommand: 'Save'
		withArgs: #()
		on: aStream.
	self class
		outputCommand: 'Quit'
		withArgs: #()
		on: aStream
]

{ #category : #accessing }
GtExcelFile >> sheetNamed: aString [
	^ sheets
		detect: [ :each | each name = aString ]
		ifNone: [ sheets add: (GtExcelSheet named: aString) ]
]
