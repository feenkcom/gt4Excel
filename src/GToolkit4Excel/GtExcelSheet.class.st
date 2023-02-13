Class {
	#name : #GtExcelSheet,
	#superclass : #Object,
	#instVars : [
		'cells',
		'name'
	],
	#category : #GToolkit4Excel
}

{ #category : #'instance creation' }
GtExcelSheet class >> named: aString [
	^ self new
		name: aString;
		yourself
]

{ #category : #accessing }
GtExcelSheet >> at: aPoint putFormula: aString [
	^ (self cellAt: aPoint)
		value: aString;
		isFormula: true
]

{ #category : #accessing }
GtExcelSheet >> at: aPoint putValue: anObject [
	^ (self cellAt: aPoint)
		value: anObject;
		isFormula: false
]

{ #category : #accessing }
GtExcelSheet >> cellAt: aPoint [
	^ cells at: aPoint ifAbsentPut: [ GtExcelCell new ]
]

{ #category : #accessing }
GtExcelSheet >> initialize [
	super initialize.
	cells := Dictionary new
]

{ #category : #accessing }
GtExcelSheet >> name [
	^ name
]

{ #category : #accessing }
GtExcelSheet >> name: aString [
	name := aString
]

{ #category : #accessing }
GtExcelSheet >> saveOn: aStream [
	GtExcelFile
		outputCommand: 'CreateSheet'
		withArgs: {self name}
		on: aStream.
	cells keysAndValuesDo: [ :point :cell | cell saveAt: point on: aStream ].
	GtExcelFile
		outputCommand: 'AdjustWidthToContents'
		withArgs: #()
		on: aStream.
	GtExcelFile
		outputCommand: 'AdjustHeightToContents'
		withArgs: #()
		on: aStream.
	GtExcelFile
		outputCommand: 'EndSheet'
		withArgs: #()
		on: aStream
]
