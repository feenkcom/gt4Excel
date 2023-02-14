Class {
	#name : #GtExcelCell,
	#superclass : #Object,
	#instVars : [
		'value',
		'isFormula',
		'isBold',
		'foreground'
	],
	#category : #GToolkit4Excel
}

{ #category : #accessing }
GtExcelCell >> bold [
	isBold := true
]

{ #category : #accessing }
GtExcelCell >> columnFor: anInteger [
	^ anInteger >= 26
		ifTrue: [ (self columnFor: anInteger // 26) , (self columnFor: anInteger \\ 26) ]
		ifFalse: [ (anInteger + 10 printStringRadix: 36) last asString ]
]

{ #category : #accessing }
GtExcelCell >> foreground [
	^ foreground
]

{ #category : #accessing }
GtExcelCell >> foreground: aColor [
	foreground := aColor
]

{ #category : #accessing }
GtExcelCell >> initialize [
	super initialize.
	isFormula := false.
	isBold := false
]

{ #category : #accessing }
GtExcelCell >> isFormula [
	^ isFormula
]

{ #category : #accessing }
GtExcelCell >> isFormula: anObject [
	isFormula := anObject
]

{ #category : #accessing }
GtExcelCell >> saveAt: aPoint on: aStream [
	| command |
	value isNil ifTrue: [ ^ self ].
	command := isFormula
			ifTrue: [ 'SetFormula' ]
			ifFalse: [ value isNumber ifTrue: [ 'SetNumber' ] ifFalse: [ 'SetText' ] ].
	GtExcelFile
		outputCommand: command
		withArgs: {(self columnFor: aPoint x - 1) , aPoint y printString.
				value}
		on: aStream.
	isBold
		ifTrue: [ GtExcelFile
				outputCommand: 'Bold'
				withArgs: {(self columnFor: aPoint x - 1) , aPoint y printString}
				on: aStream ].
	foreground
		ifNotNil: [ GtExcelFile
				outputCommand: 'ForegroundColor'
				withArgs: {(self columnFor: aPoint x - 1) , aPoint y printString.
						foreground asRGBInteger}
				on: aStream ]
]

{ #category : #accessing }
GtExcelCell >> value [
	^ value
]

{ #category : #accessing }
GtExcelCell >> value: anObject [
	value := anObject
]
