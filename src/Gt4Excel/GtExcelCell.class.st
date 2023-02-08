Class {
	#name : #GtExcelCell,
	#superclass : #Object,
	#instVars : [
		'value',
		'isFormula'
	],
	#category : #Gt4Excel
}

{ #category : #accessing }
GtExcelCell >> columnFor: anInteger [
	^ anInteger >= 26
		ifTrue: [ (self columnFor: anInteger // 26) , (self columnFor: anInteger \\ 26) ]
		ifFalse: [ (anInteger + 10 printStringRadix: 36) last asString ]
]

{ #category : #accessing }
GtExcelCell >> initialize [
	super initialize.
	isFormula := false
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

	^ GtExcelFile
		outputCommand: command
		withArgs: {value.
				self columnFor: aPoint x - 1.
				aPoint y}
		on: aStream
]

{ #category : #accessing }
GtExcelCell >> value [
	^ value
]

{ #category : #accessing }
GtExcelCell >> value: anObject [
	value := anObject
]
