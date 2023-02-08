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
	#category : #Gt4Excel
}

{ #category : #accessing }
GtExcelFile class >> generatorProgram [
	^ generatorProgram
		ifNil: [ 'ExcelGenerator' ]
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

{ #category : #accessing }
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
	filename := aString
]

{ #category : #accessing }
GtExcelFile >> initialize [
	super initialize.
	sheets := OrderedCollection new
]

{ #category : #accessing }
GtExcelFile >> save [
	| process |
	process := (GtExternalProcessBuilder new: self class generatorProgram)
			pipeStdin;
			spawn.
	self saveOn: process stdin
]

{ #category : #accessing }
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
