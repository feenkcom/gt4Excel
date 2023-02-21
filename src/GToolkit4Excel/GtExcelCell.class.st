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
GtExcelCell class >> columnFor: anInteger [
	^ anInteger >= 26
		ifTrue: [ (self columnFor: anInteger // 26) , (self columnFor: anInteger \\ 26) ]
		ifFalse: [ (anInteger + 10 printStringRadix: 36) last asString ]
]

{ #category : #accessing }
GtExcelCell >> bold [
	isBold := true
]

{ #category : #accessing }
GtExcelCell >> exportValue [
	^ (value isString and: [ value notEmpty and: [ value first = $' ] ])
		ifTrue: [ '''' , value ]
		ifFalse: [ value ]
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
GtExcelCell >> gtValueFor: aView [
	<gtView>
	^ aView forward
		title: 'Value';
		object: [ value ];
		view: #gtLiveFor:
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

{ #category : #printing }
GtExcelCell >> printOn: aStream [
	aStream << value asString
]

{ #category : #accessing }
GtExcelCell >> saveAt: aPoint on: aStream [
	| command location |
	value isNil ifTrue: [ ^ self ].
	command := isFormula
			ifTrue: [ 'SetFormula' ]
			ifFalse: [ value isNumber ifTrue: [ 'SetNumber' ] ifFalse: [ 'SetText' ] ].
	location := (self class columnFor: aPoint x - 1) , aPoint y printString.
	GtExcelFile
		outputCommand: command
		withArgs: {location.
				self exportValue}
		on: aStream.
	isBold
		ifTrue: [ GtExcelFile
				outputCommand: 'Bold'
				withArgs: {location}
				on: aStream ].
	foreground
		ifNotNil: [ GtExcelFile
				outputCommand: 'ForegroundColor'
				withArgs: {location.
						foreground asRGBInteger}
				on: aStream ]
]

{ #category : #accessing }
GtExcelCell >> text [
	| text |
	text := (value ifNil: [ '' ]) asString asRopedText.
	isBold ifTrue: [ text bold ].
	foreground ifNotNil: [ text foreground: foreground ].
	^ text
]

{ #category : #accessing }
GtExcelCell >> value [
	^ value
]

{ #category : #accessing }
GtExcelCell >> value: anObject [
	value := anObject
]
