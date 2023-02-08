Class {
	#name : #BaselineOfGt4Excel,
	#superclass : #BaselineOf,
	#category : #Gt4Excel
}

{ #category : #accessing }
BaselineOfGt4Excel >> baseline: spec [
	<baseline>
	spec for: #common do: [ spec package: 'Gt4Excel' ]
]
