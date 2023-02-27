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
GtExcelSheet >> gtSheetFor: aView [
	<gtView>
	| view rowCount columnCount |
	view := aView columnedList.
	cells isEmpty
		ifTrue: [ rowCount := 1.
			columnCount := 1 ]
		ifFalse: [ rowCount := (cells keys max: [ :each | each y ]) max: 1.
			columnCount := (cells keys max: [ :each | each x ]) max: 1 ].
	view
		title: 'Sheet';
		items: [ 1 to: rowCount ].
	view
		column: ''
		text: [ :row | 
			row printString asRopedText
				foreground: BrGlamorousColors disabledHeaderTextColor ]
		width: 25.
	1
		to: columnCount
		do: [ :x | 
			view
				column: (GtExcelCell columnFor: x - 1)
				text: [ :y | 
					cells
						at: x @ y
						ifPresent: [ :cell | cell text ]
						ifAbsent: [ '' asRopedText ] ] ].
	view
		transformation: [ :row | (1 to: columnCount) collect: [ :col | cells at: col @ row ifAbsent: [  ] ] ].
	^ view
]

{ #category : #accessing }
GtExcelSheet >> initialize [
	super initialize.
	cells := Dictionary new
]

{ #category : #accessing }
GtExcelSheet >> maxColumn [
	^ cells keys inject: 0 into: [ :max :each | max max: each x ]
]

{ #category : #accessing }
GtExcelSheet >> maxRow [
	^ cells keys inject: 0 into: [ :max :each | max max: each y ]
]

{ #category : #accessing }
GtExcelSheet >> name [
	^ name
]

{ #category : #accessing }
GtExcelSheet >> name: aString [
	name := aString
]

{ #category : #printing }
GtExcelSheet >> printOn: aStream [
	aStream << name
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
